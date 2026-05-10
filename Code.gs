const PROP_CONFIG = 'JEOPARDY_CONFIG';
const PROP_STATE = 'JEOPARDY_STATE';

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
    .setWidth(360);
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

  PropertiesService.getDocumentProperties().setProperty(PROP_CONFIG, JSON.stringify(config));
  PropertiesService.getDocumentProperties().setProperty(PROP_STATE, JSON.stringify({
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  }));

  const pres = SlidesApp.getActivePresentation();
  clearDeck_(pres);

  const titleSlide = pres.getSlides()[0];
  buildTitleSlide_(pres, titleSlide, config);

  const boardSlide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const questionMap = {};

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
  const scoringTeam = payload.passToIndex !== '' && payload.passToIndex !== null
    ? Number(payload.passToIndex)
    : Number(payload.teamIndex);

  state.scores[scoringTeam] += delta;

  if (payload.questionKey && !state.used.includes(payload.questionKey)) {
    state.used.push(payload.questionKey);
    removeBoardCell_(payload.questionKey);
  }

  state.currentTeamIndex = (Number(payload.teamIndex) + 1) % config.teams.length;

  props.setProperty(PROP_STATE, JSON.stringify(state));
  updateBoardScores_();

  return getStateForSidebar();
}

function markQuestionUsed(questionKey) {
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(PROP_STATE));
  if (!state.used.includes(questionKey)) state.used.push(questionKey);
  props.setProperty(PROP_STATE, JSON.stringify(state));
  removeBoardCell_(questionKey);
  return getStateForSidebar();
}

function resetGameState() {
  const props = PropertiesService.getDocumentProperties();
  const config = JSON.parse(props.getProperty(PROP_CONFIG) || 'null');
  if (!config) throw new Error('No game config found.');

  props.setProperty(PROP_STATE, JSON.stringify({
    scores: config.teams.map(() => 0),
    used: [],
    currentTeamIndex: 0
  }));

  buildGameFromConfig(JSON.stringify(config));
}

function validateConfig_(config) {
  if (!config.title) throw new Error('Missing title.');
  if (!Array.isArray(config.teams) || config.teams.length < 1) throw new Error('Add at least one team.');
  if (!Array.isArray(config.categories) || config.categories.length !== 6) throw new Error('You need exactly 6 categories.');
  config.categories.forEach(cat => {
    if (!cat.name) throw new Error('Every category needs a name.');
    if (!Array.isArray(cat.questions) || cat.questions.length !== 5) {
      throw new Error(`Category "${cat.name}" needs exactly 5 questions.`);
    }
  });
}

function clearDeck_(pres) {
  const slides = pres.getSlides();
  for (let i = slides.length - 1; i > 0; i--) slides[i].remove();
  slides[0].getPageElements().forEach(el => el.remove());
}

function buildTitleSlide_(pres, slide, config) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  slide.getBackground().setSolidFill('#07192f');

  const title = slide.insertTextBox(config.title, 60, 120, w - 120, 100);
  styleText_(title, 44, '#ffffff', true);

  const subtitle = slide.insertTextBox('Educational Jeopardy', 60, 230, w - 120, 50);
  styleText_(subtitle, 24, '#f7c948', false);

  const start = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, w / 2 - 90, 310, 180, 45);
  start.getText().setText('START GAME');
  styleBox_(start, '#f7c948', '#07192f', 18, true);
  start.setLinkSlide(1);
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

  const status = slide.insertTextBox('', margin, statusY, gridW, 26);
  status.setTitle('BOARD_STATUS');
  styleText_(status, 14, '#ffffff', true);

  config.teams.forEach((team, i) => {
    const boxW = gridW / config.teams.length - 6;
    const score = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, margin + i * (gridW / config.teams.length), scoreY, boxW, 34);
    score.setTitle(`SCORE::${i}`);
    score.getText().setText(`${team}: 0`);
    styleBox_(score, '#102a43', '#ffffff', 13, true);
  });

  config.categories.forEach((cat, c) => {
    const catBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + c * colW, gridY, colW - 4, rowH - 4);
    catBox.getText().setText(cat.name);
    styleBox_(catBox, '#0b3d91', '#ffffff', 13, true);

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
      cell.getText().setText(String(q.value));
      styleBox_(cell, '#123c69', '#f7c948', 22, true);
      cell.setLinkSlide(questionMap[key]);
    });
  });
}

function buildQuestionSlide_(pres, slide, answerSlide, boardSlide, categoryName, q) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  slide.getBackground().setSolidFill('#07192f');

  const header = slide.insertTextBox(`${categoryName} — ${q.value}`, 40, 25, w - 80, 45);
  styleText_(header, 24, '#f7c948', true);

  const question = slide.insertTextBox(q.question, 60, 115, w - 120, 160);
  styleText_(question, 30, '#ffffff', true);

  const answerBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, w / 2 - 105, h - 85, 210, 42);
  answerBtn.getText().setText('SHOW ANSWER');
  styleBox_(answerBtn, '#f7c948', '#07192f', 15, true);
  answerBtn.setLinkSlide(answerSlide);

  const boardBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 35, h - 85, 140, 42);
  boardBtn.getText().setText('BOARD');
  styleBox_(boardBtn, '#102a43', '#ffffff', 14, true);
  boardBtn.setLinkSlide(boardSlide);
}

function buildAnswerSlide_(pres, slide, boardSlide, categoryName, q) {
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();

  slide.getBackground().setSolidFill('#07192f');

  const header = slide.insertTextBox(`${categoryName} — Answer`, 40, 25, w - 80, 45);
  styleText_(header, 24, '#f7c948', true);

  const answer = slide.insertTextBox(q.answer, 60, 130, w - 120, 140);
  styleText_(answer, 34, '#ffffff', true);

  const boardBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, w / 2 - 100, h - 85, 200, 42);
  boardBtn.getText().setText('BACK TO BOARD');
  styleBox_(boardBtn, '#f7c948', '#07192f', 14, true);
  boardBtn.setLinkSlide(boardSlide);
}

function removeBoardCell_(questionKey) {
  const pres = SlidesApp.getActivePresentation();
  const board = pres.getSlides()[1];

  board.getPageElements().forEach(el => {
    try {
      if (el.asShape && el.asShape().getTitle() === `CELL::${questionKey}`) {
        el.remove();
      }
    } catch (e) {}
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
      const title = shape.getTitle();

      if (title && title.startsWith('SCORE::')) {
        const i = Number(title.split('::')[1]);
        shape.getText().setText(`${config.teams[i]}: ${state.scores[i]}`);
      }

      if (title === 'BOARD_STATUS') {
        shape.getText().setText(`Current turn: ${config.teams[state.currentTeamIndex]}`);
      }
    } catch (e) {}
  });
}

function styleText_(shape, size, color, bold) {
  const text = shape.getText();
  text.getTextStyle()
    .setFontSize(size)
    .setForegroundColor(color)
    .setBold(Boolean(bold));
  text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function styleBox_(shape, fill, textColor, fontSize, bold) {
  shape.getFill().setSolidFill(fill);
  shape.getLine().getFill().setSolidFill('#ffffff');
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  styleText_(shape, fontSize, textColor, bold);
}
