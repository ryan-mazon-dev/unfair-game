const UG_CONFIG = 'UNFAIR_GAME_CONFIG';
const UG_STATE = 'UNFAIR_GAME_STATE';

/**
 * Unfair Game Controller for Google Slides
 *
 * Designed for an imported Google Slides version of the "THE UNFAIR GAME" template:
 * - Home/board slide has numbered tiles 1-20.
 * - Each question slide contains "QUESTION N", question text, and a "REVEAL ANSWER" button.
 * - Each answer slide contains "ANSWER N", answer text, points, "CHANCE", and "RETURN TO HOME".
 *
 * Google Slides limitation:
 * Shape clicks in presentation mode can navigate to slides, but cannot run Apps Script.
 * For a single-screen workflow, use the modal controller/player.
 */

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Unfair Game')
    .addItem('Open Unfair Game Controller', 'showUnfairGameController')
    .addItem('Reset Game Scores / Used Tiles', 'resetUnfairGame')
    .addToUi();
}

function showUnfairGameController() {
  const html = HtmlService.createHtmlOutputFromFile('UnfairGameController')
    .setWidth(1240)
    .setHeight(760);
  SlidesApp.getUi().showModalDialog(html, 'Unfair Game Controller');
}

function getUnfairGameData() {
  const props = PropertiesService.getDocumentProperties();
  const savedConfig = JSON.parse(props.getProperty(UG_CONFIG) || 'null');
  const savedState = JSON.parse(props.getProperty(UG_STATE) || 'null');

  if (savedConfig) {
    return {
      config: savedConfig,
      state: savedState || null,
      hasSavedConfig: true,
      configSource: 'saved'
    };
  }

  // No saved setup yet: prefill the presenter form from the existing slides.
  // This lets the uploaded template's current questions/answers/points become the defaults.
  let deckConfig = null;
  try {
    deckConfig = extractUnfairConfigFromSlides_();
  } catch (e) {
    deckConfig = null;
  }

  return {
    config: deckConfig || getDefaultUnfairConfig_(),
    state: null,
    hasSavedConfig: false,
    configSource: deckConfig ? 'slides' : 'default'
  };
}

function diagnoseUnfairGameLayout() {
  const pres = SlidesApp.getActivePresentation();
  const slides = pres.getSlides();
  const layout = detectUnfairLayout_(pres);

  return {
    slideCount: slides.length,
    boardSlideIndex: layout.boardSlide ? slides.indexOf(layout.boardSlide) + 1 : null,
    questionCount: Object.keys(layout.questionSlides).length,
    answerCount: Object.keys(layout.answerSlides).length,
    question1SlideIndex: layout.questionSlides[1] ? slides.indexOf(layout.questionSlides[1]) + 1 : null,
    answer1SlideIndex: layout.answerSlides[1] ? slides.indexOf(layout.answerSlides[1]) + 1 : null
  };
}

function loadUnfairGameSetupFromSlides() {
  const config = extractUnfairConfigFromSlides_();

  const state = {
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  };

  const props = PropertiesService.getDocumentProperties();
  props.setProperty(UG_CONFIG, JSON.stringify(config));
  props.setProperty(UG_STATE, JSON.stringify(state));

  updateBoardScoreBadges_(config, state);

  return {
    config,
    state,
    hasSavedConfig: true,
    configSource: 'slides',
    message: 'Loaded setup from the current slides.'
  };
}

function extractUnfairConfigFromSlides_() {
  const pres = SlidesApp.getActivePresentation();
  const layout = detectUnfairLayout_(pres);

  if (!layout.boardSlide) {
    throw new Error('Could not find the board slide.');
  }

  const title = extractUnfairTitleFromSlides_(pres) || 'THE UNFAIR GAME!';
  const existingTeams = extractTeamsFromBoardBadges_(layout.boardSlide);

  const defaultPoints = [10, -5, 5, 1, 10, 8, 3, 10, 10, 2, 5, 7, 4, 5, 1, 5, 7, 6, 10, 10];
  const questions = [];

  for (let i = 1; i <= 20; i++) {
    const qSlide = layout.questionSlides[i];
    const aSlide = layout.answerSlides[i];

    const questionText = qSlide
      ? extractMainTextFromSlide_(qSlide, {
          preferredTitle: `UG_Q_TEXT::${i}`,
          placeholderRegex: /Insert your question here/i,
          excludeRegexes: [
            new RegExp(`^QUESTION\\s+${i}$`, 'i'),
            /^REVEAL ANSWER$/i,
            /^RETURN TO HOME$/i,
            /^CHANCE$/i,
            /^THE UNFAIR GAME!?$/i
          ]
        })
      : '';

    const answerText = aSlide
      ? extractMainTextFromSlide_(aSlide, {
          preferredTitle: `UG_A_TEXT::${i}`,
          placeholderRegex: /Insert your answer here/i,
          excludeRegexes: [
            new RegExp(`^ANSWER\\s+${i}$`, 'i'),
            /^REVEAL ANSWER$/i,
            /^RETURN TO HOME$/i,
            /^CHANCE$/i,
            /^-?\\d+\\s+points?$/i,
            /^THE UNFAIR GAME!?$/i
          ]
        })
      : '';

    const points = aSlide ? extractPointsFromAnswerSlide_(aSlide, i, defaultPoints[i - 1]) : defaultPoints[i - 1];

    questions.push({
      number: i,
      question: questionText || `Question ${i}`,
      answer: answerText || `Answer ${i}`,
      points
    });
  }

  return {
    title,
    teams: existingTeams.length ? existingTeams : ['Team 1', 'Team 2'],
    questions
  };
}

function extractUnfairTitleFromSlides_(pres) {
  const slides = pres.getSlides();
  let fallback = '';

  slides.forEach(slide => {
    walkPageElements_(slide.getPageElements(), el => {
      try {
        const shape = el.asShape();
        const text = shape.getText().asString().trim();

        if (/THE\\s+UNFAIR\\s+GAME!?/i.test(text)) {
          fallback = text;
        }
      } catch (e) {}
    });
  });

  return fallback;
}

function extractTeamsFromBoardBadges_(boardSlide) {
  const teams = [];

  walkPageElements_(boardSlide.getPageElements(), el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      const text = shape.getText().asString().trim();

      if (title && title.indexOf('UG_SCORE_BADGE::') === 0) {
        const teamName = text.split(':')[0].trim();
        if (teamName && !teams.includes(teamName)) {
          teams.push(teamName);
        }
      }
    } catch (e) {}
  });

  return teams;
}

function extractMainTextFromSlide_(slide, options) {
  const candidates = [];
  let preferred = '';
  let placeholder = '';

  walkPageElements_(slide.getPageElements(), el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      const raw = shape.getText().asString();
      const text = String(raw || '').trim();

      if (!text) return;

      if (title === options.preferredTitle) {
        preferred = text;
        return;
      }

      if (options.placeholderRegex && options.placeholderRegex.test(text)) {
        placeholder = text;
        return;
      }

      const excluded = (options.excludeRegexes || []).some(rx => rx.test(text));
      if (!excluded) {
        candidates.push(text);
      }
    } catch (e) {}
  });

  if (preferred) return preferred;
  if (placeholder) return placeholder;

  // Best guess: actual question/answer text is normally the longest remaining text box.
  candidates.sort((a, b) => b.length - a.length);
  return candidates[0] || '';
}

function extractPointsFromAnswerSlide_(slide, questionNumber, fallback) {
  let points = null;

  walkPageElements_(slide.getPageElements(), el => {
    if (points !== null) return;

    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      const text = shape.getText().asString().trim();

      if (title === `UG_POINTS::${questionNumber}` || /^-?\\d+\\s+points?$/i.test(text)) {
        const match = text.match(/-?\\d+/);
        if (match) {
          points = Number(match[0]);
        }
      }
    } catch (e) {}
  });

  return points !== null && !Number.isNaN(points) ? points : Number(fallback || 0);
}

function saveUnfairGameSetup(payload) {
  const config = normalizeUnfairConfig_(payload);

  const state = {
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  };

  PropertiesService.getDocumentProperties().setProperty(UG_CONFIG, JSON.stringify(config));
  PropertiesService.getDocumentProperties().setProperty(UG_STATE, JSON.stringify(state));

  updateUnfairSlidesFromConfig_(config);

  return {
    config,
    state,
    hasSavedConfig: true,
    message: 'Setup saved. Slides updated and controller is ready.'
  };
}

function applyUnfairResult(payload) {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(UG_CONFIG) || 'null');
  const state = JSON.parse(props.getProperty(UG_STATE) || 'null');

  if (!config || !state) {
    throw new Error('No saved game setup found. Save setup first.');
  }

  const questionNumber = Number(payload.questionNumber);
  const answeringTeam = Number(payload.teamIndex || 0);
  const passToRaw = payload.passToIndex;

  const passToTeam = passToRaw !== '' && passToRaw !== null && passToRaw !== undefined
    ? Number(passToRaw)
    : null;

  const scoringTeam = passToTeam !== null ? passToTeam : answeringTeam;
  const delta = Number(payload.delta || 0);

  state.scores[scoringTeam] += delta;

  if (questionNumber && !state.used.includes(questionNumber)) {
    state.used.push(questionNumber);
    removeBoardTile_(questionNumber);
  }

  state.currentTeamIndex = (answeringTeam + 1) % config.teams.length;

  props.setProperty(UG_STATE, JSON.stringify(state));
  updateBoardScoreBadges_(config, state);

  return {
    config,
    state,
    hasSavedConfig: true
  };
}

function markUnfairQuestionUsed(questionNumber) {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(UG_CONFIG) || 'null');
  const state = JSON.parse(props.getProperty(UG_STATE) || 'null');

  if (!config || !state) {
    throw new Error('No saved game setup found. Save setup first.');
  }

  const n = Number(questionNumber);

  if (n && !state.used.includes(n)) {
    state.used.push(n);
    removeBoardTile_(n);
  }

  props.setProperty(UG_STATE, JSON.stringify(state));
  updateBoardScoreBadges_(config, state);

  return {
    config,
    state,
    hasSavedConfig: true
  };
}

function resetUnfairGame() {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(UG_CONFIG) || 'null');

  if (!config) {
    throw new Error('No saved game setup found. Save setup first.');
  }

  const state = {
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  };

  props.setProperty(UG_STATE, JSON.stringify(state));
  updateUnfairSlidesFromConfig_(config);

  return {
    config,
    state,
    hasSavedConfig: true,
    message: 'Scores reset and all board tiles restored.'
  };
}

function getDefaultUnfairConfig_() {
  const questions = [];
  const defaultPoints = [10, -5, 5, 1, 10, 8, 3, 10, 10, 2, 5, 7, 4, 5, 1, 5, 7, 6, 10, 10];

  for (let i = 1; i <= 20; i++) {
    questions.push({
      number: i,
      question: '',
      answer: '',
      points: defaultPoints[i - 1]
    });
  }

  return {
    title: 'THE UNFAIR GAME!',
    teams: ['Team 1', 'Team 2'],
    questions
  };
}

function normalizeUnfairConfig_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Missing setup data.');
  }

  const title = String(payload.title || 'THE UNFAIR GAME!').trim() || 'THE UNFAIR GAME!';

  let teams = Array.isArray(payload.teams) ? payload.teams : [];
  teams = teams
    .map(t => String(t || '').trim())
    .filter(Boolean);

  if (teams.length < 1) {
    throw new Error('Add at least one team.');
  }

  if (teams.length > 12) {
    throw new Error('Use 12 or fewer teams. The controller supports more, but the room will riot.');
  }

  const rawQuestions = Array.isArray(payload.questions) ? payload.questions : [];
  if (rawQuestions.length !== 20) {
    throw new Error('This Unfair Game layout needs exactly 20 questions.');
  }

  const questions = rawQuestions.map((q, idx) => {
    const number = idx + 1;
    const question = String(q.question || '').trim() || `Question ${number}`;
    const answer = String(q.answer || '').trim() || `Answer ${number}`;
    const points = Number(q.points || 0);

    if (Number.isNaN(points)) {
      throw new Error(`Question ${number} has an invalid point value.`);
    }

    return { number, question, answer, points };
  });

  return { title, teams, questions };
}

function updateUnfairSlidesFromConfig_(config) {
  const pres = SlidesApp.getActivePresentation();
  const layout = detectUnfairLayout_(pres);

  if (!layout.boardSlide) {
    throw new Error('Could not find the board slide with tiles 1-20.');
  }

  setUnfairTitle_(pres, config.title);
  updateBoardSlide_(layout.boardSlide, layout, config);

  config.questions.forEach(q => {
    const qSlide = layout.questionSlides[q.number];
    const aSlide = layout.answerSlides[q.number];

    if (!qSlide) {
      throw new Error(`Could not find QUESTION ${q.number} slide.`);
    }

    if (!aSlide) {
      throw new Error(`Could not find ANSWER ${q.number} slide.`);
    }

    updateQuestionSlide_(qSlide, aSlide, q);
    updateAnswerSlide_(aSlide, layout.boardSlide, q);
  });

  updateBoardScoreBadges_(config, {
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  });
}

function detectUnfairLayout_(pres) {
  const slides = pres.getSlides();
  const questionSlides = {};
  const answerSlides = {};
  let boardSlide = null;
  let bestBoardCandidate = null;
  let bestBoardScore = -1;

  slides.forEach((slide, index) => {
    const texts = getSlideTexts_(slide);
    const blob = texts.join('\n');
    const uniqueNumbers = getUniqueBoardNumbersFromTexts_(texts);
    const hasUnfairTitle = /THE\s+UNFAIR\s+GAME!?/i.test(blob);

    // Strong board detection:
    // The template board has THE UNFAIR GAME plus numbers 1-20.
    // After PPTX import, numbers may be grouped or oddly formatted, so count tokens,
    // not just exact text boxes.
    const boardScore = uniqueNumbers.length + (hasUnfairTitle ? 10 : 0);

    if (boardScore > bestBoardScore) {
      bestBoardScore = boardScore;
      bestBoardCandidate = slide;
    }

    if (!boardSlide && uniqueNumbers.length >= 15) {
      boardSlide = slide;
    }

    texts.forEach(t => {
      let match = String(t).match(/\bQUESTION\s+(\d+)\b/i);
      if (match) {
        questionSlides[Number(match[1])] = slide;
      }

      match = String(t).match(/\bANSWER\s+(\d+)\b/i);
      if (match) {
        answerSlides[Number(match[1])] = slide;
      }
    });
  });

  // Fallback for this exact template:
  // In the uploaded PowerPoint, there are blank/transition slides before the board,
  // and QUESTION 1 comes immediately after the board. If number detection fails
  // because Google imported the board as grouped objects, use the slide before QUESTION 1.
  if (!boardSlide && questionSlides[1]) {
    const q1Index = slides.indexOf(questionSlides[1]);
    if (q1Index > 0) {
      boardSlide = slides[q1Index - 1];
    }
  }

  // Final fallback: use the highest-scoring slide candidate.
  if (!boardSlide && bestBoardCandidate && bestBoardScore > 0) {
    boardSlide = bestBoardCandidate;
  }

  return { boardSlide, questionSlides, answerSlides };
}

function getSlideTexts_(slide) {
  const texts = [];

  walkPageElements_(slide.getPageElements(), el => {
    const text = getElementText_(el);
    if (text !== null && String(text).trim() !== '') {
      texts.push(String(text).trim());
    }
  });

  return texts;
}

function getUniqueBoardNumbersFromTexts_(texts) {
  const found = {};

  texts.forEach(text => {
    const matches = String(text).match(/\b([1-9]|1[0-9]|20)\b/g) || [];
    matches.forEach(m => {
      const n = Number(m);
      if (n >= 1 && n <= 20) {
        found[n] = true;
      }
    });
  });

  return Object.keys(found).map(Number);
}

function walkPageElements_(elements, callback) {
  elements.forEach(el => {
    try {
      callback(el);
    } catch (e) {}

    // PPTX imports sometimes wrap board tiles/buttons inside groups.
    // Recursing into groups makes detection and linking much more reliable.
    try {
      const group = el.asGroup();
      if (group) {
        walkPageElements_(group.getChildren(), callback);
      }
    } catch (e) {}
  });
}

function getElementText_(el) {
  try {
    const shape = el.asShape();
    const text = shape.getText().asString();
    return String(text || '').trim();
  } catch (e) {
    return null;
  }
}

function setElementText_(el, value) {
  try {
    const shape = el.asShape();
    shape.getText().setText(String(value));
    return true;
  } catch (e) {
    return false;
  }
}

function setShapeText_(shape, value) {
  try {
    shape.getText().setText(String(value));
    return true;
  } catch (e) {
    return false;
  }
}

function setUnfairTitle_(pres, title) {
  pres.getSlides().forEach(slide => {
    slide.getPageElements().forEach(el => {
      try {
        const shape = el.asShape();
        const text = shape.getText().asString().trim();
        if (/^THE\s+UNFAIR\s+GAME!?$/i.test(text)) {
          setShapeText_(shape, title);
        }
      } catch (e) {}
    });
  });
}

function updateBoardSlide_(boardSlide, layout, config) {
  walkPageElements_(boardSlide.getPageElements(), el => {
    try {
      const shape = el.asShape();
      const text = shape.getText().asString().trim();

      if (/^THE\s+UNFAIR\s+GAME!?$/i.test(text)) {
        setShapeText_(shape, config.title);
      }

      const n = Number(text);
      if (n >= 1 && n <= 20 && String(n) === text) {
        shape.setTitle(`UG_TILE::${n}`);
        const qSlide = layout.questionSlides[n];
        if (qSlide) {
          shape.setLinkSlide(qSlide);
        }
      }
    } catch (e) {}
  });
}

function updateQuestionSlide_(qSlide, answerSlide, q) {
  qSlide.getPageElements().forEach(el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      const text = shape.getText().asString().trim();

      if (title === `UG_Q_TEXT::${q.number}` || /Insert your question here/i.test(text)) {
        shape.setTitle(`UG_Q_TEXT::${q.number}`);
        setShapeText_(shape, q.question);
        return;
      }

      if (new RegExp(`^QUESTION\\s+${q.number}$`, 'i').test(text)) {
        shape.setTitle(`UG_Q_TITLE::${q.number}`);
        setShapeText_(shape, `QUESTION ${q.number}`);
        return;
      }

      if (/REVEAL ANSWER/i.test(text)) {
        shape.setTitle(`UG_REVEAL::${q.number}`);
        shape.setLinkSlide(answerSlide);
        return;
      }
    } catch (e) {}
  });
}

function updateAnswerSlide_(aSlide, boardSlide, q) {
  aSlide.getPageElements().forEach(el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      const text = shape.getText().asString().trim();

      if (title === `UG_A_TEXT::${q.number}` || /Insert your answer here/i.test(text)) {
        shape.setTitle(`UG_A_TEXT::${q.number}`);
        setShapeText_(shape, q.answer);
        return;
      }

      if (new RegExp(`^ANSWER\\s+${q.number}$`, 'i').test(text)) {
        shape.setTitle(`UG_A_TITLE::${q.number}`);
        setShapeText_(shape, `ANSWER ${q.number}`);
        return;
      }

      if (title === `UG_POINTS::${q.number}` || /^-?\d+\s+points?$/i.test(text)) {
        shape.setTitle(`UG_POINTS::${q.number}`);
        setShapeText_(shape, formatPoints_(q.points));
        return;
      }

      if (/RETURN TO HOME/i.test(text)) {
        shape.setTitle(`UG_RETURN::${q.number}`);
        shape.setLinkSlide(boardSlide);
        return;
      }
    } catch (e) {}
  });
}

function removeBoardTile_(questionNumber) {
  const pres = SlidesApp.getActivePresentation();
  const layout = detectUnfairLayout_(pres);

  if (!layout.boardSlide) return;

  walkPageElements_(layout.boardSlide.getPageElements(), el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      const text = shape.getText().asString().trim();
      const n = Number(questionNumber);

      if (title === `UG_TILE::${n}` || text === String(n)) {
        el.remove();
      }
    } catch (e) {}
  });
}

function updateBoardScoreBadges_(config, state) {
  const pres = SlidesApp.getActivePresentation();
  const layout = detectUnfairLayout_(pres);

  if (!layout.boardSlide) return;

  // Remove old score/status badges created by this script.
  walkPageElements_(layout.boardSlide.getPageElements(), el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';
      if (title && title.indexOf('UG_SCORE_BADGE::') === 0) {
        el.remove();
      }
      if (title === 'UG_TURN_BADGE') {
        el.remove();
      }
    } catch (e) {}
  });

  const presWidth = pres.getPageWidth();
  const y = 10;
  const h = 26;
  const gap = 6;
  const maxBadges = Math.max(1, config.teams.length);
  const badgeW = Math.max(70, Math.min(160, (presWidth - 20 - (gap * (maxBadges - 1))) / maxBadges));

  config.teams.forEach((team, i) => {
    const x = 10 + i * (badgeW + gap);
    const badge = layout.boardSlide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, badgeW, h);
    badge.setTitle(`UG_SCORE_BADGE::${i}`);
    badge.getFill().setSolidFill(i === state.currentTeamIndex ? '#f7c948' : '#102a43');
    try {
      badge.getBorder().getLineFill().setSolidFill('#ffffff');
    } catch (e) {}
    badge.getText().setText(`${team}: ${state.scores[i]}`);
    try {
      badge.getText().getTextStyle()
        .setFontSize(10)
        .setForegroundColor(i === state.currentTeamIndex ? '#07192f' : '#ffffff')
        .setBold(true);
    } catch (e) {}
    try {
      badge.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      badge.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    } catch (e) {}
  });
}

function formatPoints_(points) {
  const n = Number(points);
  const word = Math.abs(n) === 1 ? 'point' : 'points';
  return `${n} ${word}`;
}
