const PROP_CONFIG = 'JEOPARDY_CONFIG';
const PROP_STATE = 'JEOPARDY_STATE';

/**
 * Google Slides Jeopardy Game Builder
 *
 * Important Google Slides limitation:
 * A shape click in presentation mode can navigate to another slide,
 * but it cannot directly run Apps Script. This script uses:
 *
 * 1. Slide links for board/question/answer navigation.
 * 2. A sidebar controller for scoring and persistent cell removal.
 */

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Jeopardy')
    .addItem('Open Controller', 'showController')
    .addItem('Reset Scores / Board', 'resetGameState')
    .addToUi();
}

function showController() {
  const html = HtmlService.createHtmlOutputFromFile('Controller')
    .setTitle('Jeopardy Controller')
    .setWidth(380);
  SlidesApp.getUi().showSidebar(html);
}

function getSampleConfig() {
  return JSON.stringify({
    title: 'Classroom Jeopardy',
    teams: ['Team 1', 'Team 2', 'Team 3'],
    categories: [
      {
        name: 'History',
        questions: [
          { value: 100, question: 'Who was the first U.S. president?', answer: 'George Washington' },
          { value: 200, question: 'What year did World War II end?', answer: '1945' },
          { value: 300, question: 'What empire was ruled from Constantinople?', answer: 'The Byzantine Empire' },
          { value: 400, question: 'Who wrote the Declaration of Independence?', answer: 'Thomas Jefferson' },
          { value: 500, question: 'What ancient city was buried by Mount Vesuvius?', answer: 'Pompeii' }
        ]
      },
      {
        name: 'Science',
        questions: [
          { value: 100, question: 'What planet is known as the Red Planet?', answer: 'Mars' },
          { value: 200, question: 'What gas do plants absorb?', answer: 'Carbon dioxide' },
          { value: 300, question: 'What is H2O?', answer: 'Water' },
          { value: 400, question: 'What force keeps us on Earth?', answer: 'Gravity' },
          { value: 500, question: 'What particle has a negative charge?', answer: 'Electron' }
        ]
      },
      {
        name: 'Math',
        questions: [
          { value: 100, question: 'What is 8 × 7?', answer: '56' },
          { value: 200, question: 'What is the square root of 81?', answer: '9' },
          { value: 300, question: 'What is 15% of 200?', answer: '30' },
          { value: 400, question: 'What is the area of a triangle formula?', answer: '1/2 × base × height' },
          { value: 500, question: 'What is 2 to the 5th power?', answer: '32' }
        ]
      },
      {
        name: 'Literature',
        questions: [
          { value: 100, question: 'Who wrote Romeo and Juliet?', answer: 'William Shakespeare' },
          { value: 200, question: 'What is the main character called?', answer: 'The protagonist' },
          { value: 300, question: 'What is a comparison using “like” or “as”?', answer: 'Simile' },
          { value: 400, question: 'Who wrote The Hobbit?', answer: 'J.R.R. Tolkien' },
          { value: 500, question: 'What is dramatic irony?', answer: 'When the audience knows something a character does not' }
        ]
      },
      {
        name: 'Geography',
        questions: [
          { value: 100, question: 'What is the largest ocean?', answer: 'Pacific Ocean' },
          { value: 200, question: 'What continent is Egypt in?', answer: 'Africa' },
          { value: 300, question: 'What is the capital of France?', answer: 'Paris' },
          { value: 400, question: 'What river runs through London?', answer: 'The Thames' },
          { value: 500, question: 'What country has the most people?', answer: 'India' }
        ]
      },
      {
        name: 'Wildcard',
        questions: [
          { value: 100, question: 'How many sides does a hexagon have?', answer: '6' },
          { value: 200, question: 'What color do you get from blue and yellow?', answer: 'Green' },
          { value: 300, question: 'What is the opposite of “ancient”?', answer: 'Modern' },
          { value: 400, question: 'What instrument has keys, pedals, and strings?', answer: 'Piano' },
          { value: 500, question: 'What is the fastest land animal?', answer: 'Cheetah' }
        ]
      }
    ]
  }, null, 2);
}

function buildGameFromConfig(configText) {
  const config = JSON.parse(configText);
  validateConfig_(config);

  const props = PropertiesService.getDocumentProperties();
  props.setProperty(PROP_CONFIG, JSON.stringify(config));
  props.setProperty(PROP_STATE, JSON.stringify({
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  }));

  const pres = SlidesApp.getActivePresentation();
  clearDeck_(pres);

  const titleSlide = pres.getSlides()[0];
  const boardSlide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const questionMap = {};

  buildTitleSlide_(pres, titleSlide, config, boardSlide);

  config.categories.forEach((cat, c) => {
    cat.questions.forEach((q, r) => {
      const qSlide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const aSlide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

      buildQuestionSlide_(pres, qSlide, aSlide, boardSlide, cat.name, q);
      buildAnswerSlide_(pres, aSlide, boardSlide, cat.name, q);

      questionMap[`C${c}Q${r}`] = qSlide;
    });
  });

  buildBoardSlide_(pres, boardSlide, config, questionMap);
  updateBoardScores_();

  return 'Game built. Use Present for the deck, and keep the Jeopardy Controller open for scoring/removing cells.';
}

function getStateForSidebar() {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(PROP_CONFIG) || 'null');
  const state = JSON.parse(props.getProperty(PROP_STATE) || 'null');
  return { config, state };
}

function applyResult(payload) {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(PROP_CONFIG));
  const state = JSON.parse(props.getProperty(PROP_STATE));

  const delta = Number(payload.delta || 0);
  const answeringTeam = Number(payload.teamIndex);
  const passedToTeam = payload.passToIndex !== '' && payload.passToIndex !== null
    ? Number(payload.passToIndex)
    : null;

  const scoringTeam = passedToTeam !== null ? passedToTeam : answeringTeam;
  state.scores[scoringTeam] += delta;

  if (payload.questionKey && !state.used.includes(payload.questionKey)) {
    state.used.push(payload.questionKey);
    removeBoardCell_(payload.questionKey);
  }

  state.currentTeamIndex = (answeringTeam + 1) % config.teams.length;

  props.setProperty(PROP_STATE, JSON.stringify(state));
  updateBoardScores_();

  return getStateForSidebar();
}

function markQuestionUsed(questionKey) {
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(PROP_STATE));

  if (questionKey && !state.used.includes(questionKey)) {
    state.used.push(questionKey);
  }

  props.setProperty(PROP_STATE, JSON.stringify(state));
  removeBoardCell_(questionKey);

  return getStateForSidebar();
}

function resetGameState() {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(PROP_CONFIG) || 'null');

  if (!config) {
    throw new Error('No game config found. Build the game first.');
  }

  buildGameFromConfig(JSON.stringify(config));
}

function validateConfig_(config) {
  if (!config || typeof config !== 'object') {
    throw new Error('Config must be a JSON object.');
  }

  if (!config.title) {
    throw new Error('Missing title.');
  }

  if (!Array.isArray(config.teams) || config.teams.length < 1) {
    throw new Error('Add at least one team.');
  }

  if (!Array.isArray(config.categories) || config.categories.length !== 6) {
    throw new Error('You need exactly 6 categories.');
  }

  config.categories.forEach((cat, c) => {
    if (!cat.name) {
      throw new Error(`Category ${c + 1} is missing a name.`);
    }

    if (!Array.isArray(cat.questions) || cat.questions.length !== 5) {
      throw new Error(`Category "${cat.name}" needs exactly 5 questions.`);
    }

    cat.questions.forEach((q, r) => {
      if (q.value === undefined || q.value === null) {
        throw new Error(`Question ${r + 1} in "${cat.name}" is missing a value.`);
      }
      if (!q.question) {
        throw new Error(`Question ${r + 1} in "${cat.name}" is missing question text.`);
      }
      if (!q.answer) {
        throw new Error(`Question ${r + 1} in "${cat.name}" is missing answer text.`);
      }
    });
  });
}

function clearDeck_(pres) {
  const slides = pres.getSlides();

  for (let i = slides.length - 1; i > 0; i--) {
    slides[i].remove();
  }

  const firstSlide = slides[0];
  firstSlide.getPageElements().forEach(el => {
    try {
      el.remove();
    } catch (e) {
      // Ignore stubborn placeholder artifacts.
    }
  });
}

function buildTitleSlide_(pres, slide, config, boardSlide) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  slide.getBackground().setSolidFill('#07192f');

  const title = slide.insertTextBox(String(config.title), 60, 120, w - 120, 100);
  styleTextShape_(title, String(config.title), 44, '#ffffff', true);

  const subtitleText = 'Educational Jeopardy';
  const subtitle = slide.insertTextBox(subtitleText, 60, 230, w - 120, 50);
  styleTextShape_(subtitle, subtitleText, 24, '#f7c948', false);

  const start = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, w / 2 - 90, 310, 180, 45);
  styleShapeWithText_(start, 'START GAME', '#f7c948', '#07192f', 18, true);
  start.setLinkSlide(boardSlide);
}

function buildBoardSlide_(pres, slide, config, questionMap) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  const margin = 24;
  const scoreY = 14;
  const statusY = 58;
  const gridY = 92;

  const gridW = w - margin * 2;
  const gridH = h - gridY - 18;
  const colW = gridW / 6;
  const rowH = gridH / 6;

  slide.getBackground().setSolidFill('#07192f');

  const statusText = 'Current turn: ' + config.teams[0];
  const status = slide.insertTextBox(statusText, margin, statusY, gridW, 26);
  status.setTitle('BOARD_STATUS');
  styleTextShape_(status, statusText, 14, '#ffffff', true);

  config.teams.forEach((team, i) => {
    const boxW = gridW / config.teams.length - 6;
    const score = slide.insertShape(
      SlidesApp.ShapeType.ROUND_RECTANGLE,
      margin + i * (gridW / config.teams.length),
      scoreY,
      boxW,
      34
    );

    score.setTitle(`SCORE::${i}`);
    styleShapeWithText_(score, `${team}: 0`, '#102a43', '#ffffff', 13, true);
  });

  config.categories.forEach((cat, c) => {
    const catBox = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      margin + c * colW,
      gridY,
      colW - 4,
      rowH - 4
    );

    styleShapeWithText_(catBox, cat.name, '#0b3d91', '#ffffff', 13, true);

    cat.questions.forEach((q, r) => {
      const key = `C${c}Q${r}`;

      const cell = slide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        margin + c * colW,
        gridY + (r + 1) * rowH,
        colW - 4,
        rowH - 4
      );

      cell.setTitle(`CELL::${key}`);
      styleShapeWithText_(cell, String(q.value), '#123c69', '#f7c948', 22, true);
      cell.setLinkSlide(questionMap[key]);
    });
  });
}

function buildQuestionSlide_(pres, slide, answerSlide, boardSlide, categoryName, q) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  slide.getBackground().setSolidFill('#07192f');

  const headerText = `${categoryName} — ${q.value}`;
  const header = slide.insertTextBox(headerText, 40, 25, w - 80, 45);
  styleTextShape_(header, headerText, 24, '#f7c948', true);

  const questionText = String(q.question);
  const question = slide.insertTextBox(questionText, 60, 115, w - 120, 180);
  styleTextShape_(question, questionText, 30, '#ffffff', true);

  const answerBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, w / 2 - 105, h - 85, 210, 42);
  styleShapeWithText_(answerBtn, 'SHOW ANSWER', '#f7c948', '#07192f', 15, true);
  answerBtn.setLinkSlide(answerSlide);

  const boardBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 35, h - 85, 140, 42);
  styleShapeWithText_(boardBtn, 'BOARD', '#102a43', '#ffffff', 14, true);
  boardBtn.setLinkSlide(boardSlide);
}

function buildAnswerSlide_(pres, slide, boardSlide, categoryName, q) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  slide.getBackground().setSolidFill('#07192f');

  const headerText = `${categoryName} — Answer`;
  const header = slide.insertTextBox(headerText, 40, 25, w - 80, 45);
  styleTextShape_(header, headerText, 24, '#f7c948', true);

  const answerText = String(q.answer);
  const answer = slide.insertTextBox(answerText, 60, 130, w - 120, 160);
  styleTextShape_(answer, answerText, 34, '#ffffff', true);

  const boardBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, w / 2 - 100, h - 85, 200, 42);
  styleShapeWithText_(boardBtn, 'BACK TO BOARD', '#f7c948', '#07192f', 14, true);
  boardBtn.setLinkSlide(boardSlide);
}

function removeBoardCell_(questionKey) {
  if (!questionKey) return;

  const pres = SlidesApp.getActivePresentation();
  const board = pres.getSlides()[1];

  board.getPageElements().forEach(el => {
    try {
      const shape = el.asShape();
      if (shape.getTitle && shape.getTitle() === `CELL::${questionKey}`) {
        el.remove();
      }
    } catch (e) {
      // Ignore non-shape objects.
    }
  });
}

function updateBoardScores_() {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(PROP_CONFIG));
  const state = JSON.parse(props.getProperty(PROP_STATE));

  const pres = SlidesApp.getActivePresentation();
  const board = pres.getSlides()[1];

  board.getPageElements().forEach(el => {
    try {
      const shape = el.asShape();
      const title = shape.getTitle ? shape.getTitle() : '';

      if (title && title.startsWith('SCORE::')) {
        const i = Number(title.split('::')[1]);
        const scoreText = `${config.teams[i]}: ${state.scores[i]}`;
        styleShapeWithText_(shape, scoreText, '#102a43', '#ffffff', 13, true);
      }

      if (title === 'BOARD_STATUS') {
        const statusText = `Current turn: ${config.teams[state.currentTeamIndex]}`;
        styleTextShape_(shape, statusText, 14, '#ffffff', true);
      }
    } catch (e) {
      // Ignore non-shape objects.
    }
  });
}

/**
 * Safely styles a text box.
 * This function always sets text first, then styles the fresh non-empty range.
 */
function styleTextShape_(shape, textValue, fontSize, textColor, bold) {
  const value = safeVisibleText_(textValue);

  try {
    shape.getText().setText(value);
  } catch (e) {
    return;
  }

  try {
    const range = shape.getText();
    if (range.isEmpty && range.isEmpty()) return;

    const style = range.getTextStyle();
    if (style) {
      style
        .setFontSize(fontSize)
        .setForegroundColor(textColor)
        .setBold(Boolean(bold));
    }
  } catch (e) {
    // Text styling failed, but the actual text exists. Keep going.
  }

  try {
    const paragraphStyle = shape.getText().getParagraphStyle();
    if (paragraphStyle) {
      paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
  } catch (e) {
    // Paragraph styling is cosmetic.
  }
}

/**
 * Safely styles a shape and its internal text.
 */
function styleShapeWithText_(shape, textValue, fillColor, textColor, fontSize, bold) {
  try {
    shape.getFill().setSolidFill(fillColor);
  } catch (e) {}

  try {
    const border = shape.getBorder();
    if (border) {
      border.getLineFill().setSolidFill('#ffffff');
      border.setWeight(1);
    }
  } catch (e) {}

  try {
    shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}

  styleTextShape_(shape, textValue, fontSize, textColor, bold);
}

/**
 * Apps Script gets cranky about styling empty TextRanges.
 * This guarantees at least one visible character.
 */
function safeVisibleText_(value) {
  if (value === undefined || value === null) return ' ';
  const text = String(value);
  return text.length > 0 ? text : ' ';
}
