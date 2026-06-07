function collectNewsCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keywordSheet = ss.getSheetByName('キーワード設定');
  const candidateSheet = ss.getSheetByName('投稿候補');

  if (!keywordSheet || !candidateSheet) {
    throw new Error('キーワード設定 または 投稿候補 シートが見つかりません。');
  }

  const keywordRows = keywordSheet.getDataRange().getValues();
  const candidateHeaders = candidateSheet.getRange(1, 1, 1, candidateSheet.getLastColumn()).getValues()[0];

  const col = name => candidateHeaders.indexOf(name) + 1;
  const collectedAtCol = col('取得日時');
  const sourceCol = col('情報源');
  const urlCol = col('元記事URL');
  const titleCol = col('タイトル');
  const sourceTextCol = col('元情報');
  const statusCol = col('ステータス');

  if (!collectedAtCol || !sourceCol || !urlCol || !titleCol || !sourceTextCol || !statusCol) {
    throw new Error('投稿候補シートの列名を確認してください。');
  }

  const existingUrls = getExistingUrls(candidateSheet, urlCol);

  for (let i = 1; i < keywordRows.length; i++) {
    const keyword = keywordRows[i][0];
    const target = keywordRows[i][1];
    const priority = keywordRows[i][2];

    if (!keyword) continue;
    if (target && String(target).indexOf('ニュース') === -1) continue;

    const items = fetchGoogleNewsRss(keyword);

    items.forEach(item => {
      if (existingUrls.has(item.link)) return;

      const row = candidateSheet.getLastRow() + 1;

      candidateSheet.getRange(row, collectedAtCol).setValue(new Date());
      candidateSheet.getRange(row, sourceCol).setValue(`Google News RSS / ${keyword} / ${priority || ''}`);
      candidateSheet.getRange(row, urlCol).setValue(item.link);
      candidateSheet.getRange(row, titleCol).setValue(item.title);
      candidateSheet.getRange(row, sourceTextCol).setValue(item.description || item.title);
      candidateSheet.getRange(row, statusCol).setValue('候補');

      existingUrls.add(item.link);
    });
  }
}

function fetchGoogleNewsRss(keyword) {
  const query = encodeURIComponent(keyword);
  const url = `https://news.google.com/rss/search?q=${query}&hl=ja&gl=JP&ceid=JP:ja`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 400) {
    throw new Error(response.getContentText());
  }

  const xml = XmlService.parse(response.getContentText());
  const root = xml.getRootElement();
  const channel = root.getChild('channel');
  const items = channel.getChildren('item');

  return items.slice(0, 10).map(item => ({
    title: getXmlText(item, 'title'),
    link: getXmlText(item, 'link'),
    description: stripHtml(getXmlText(item, 'description'))
  }));
}

function getXmlText(element, childName) {
  const child = element.getChild(childName);
  return child ? child.getText() : '';
}

function stripHtml(text) {
  return String(text || '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getExistingUrls(sheet, urlCol) {
  const lastRow = sheet.getLastRow();
  const urls = new Set();

  if (lastRow < 2) return urls;

  const values = sheet.getRange(2, urlCol, lastRow - 1, 1).getValues();
  values.forEach(row => {
    if (row[0]) urls.add(String(row[0]));
  });

  return urls;
}

function generateSelectedCandidatePosts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('投稿候補');
  const range = sheet.getActiveRange();
  generateCandidatePostsForRange(sheet, range);
}

function generateCandidatePostsForRange(sheet, range) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = name => headers.indexOf(name) + 1;

  const urlCol = col('元記事URL');
  const titleCol = col('タイトル');
  const sourceTextCol = col('元情報');
  const factCol = col('事実');
  const inferenceCol = col('推測');
  const questionCol = col('疑問点');
  const interestCol = col('利害関係');
  const historyCol = col('過去行動');
  const aiCol = col('AI投稿文');
  const statusCol = col('ステータス');

  if (!urlCol || !titleCol || !sourceTextCol || !factCol || !inferenceCol || !questionCol || !interestCol || !historyCol || !aiCol || !statusCol) {
    throw new Error('投稿候補シートの列名を確認してください。');
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row === 1) continue;

    const existingPost = sheet.getRange(row, aiCol).getValue();
    if (existingPost) continue;

    const data = {
      url: sheet.getRange(row, urlCol).getDisplayValue(),
      title: sheet.getRange(row, titleCol).getDisplayValue(),
      sourceText: sheet.getRange(row, sourceTextCol).getDisplayValue()
    };

    const result = createPoliticalAnalysisAndPost(data);

    sheet.getRange(row, factCol).setValue(result.fact);
    sheet.getRange(row, inferenceCol).setValue(result.inference);
    sheet.getRange(row, questionCol).setValue(result.question);
    sheet.getRange(row, interestCol).setValue(result.interest);
    sheet.getRange(row, historyCol).setValue(result.history);
    sheet.getRange(row, aiCol).setValue(result.post);
    sheet.getRange(row, statusCol).setValue('確認待ち');
  }
}

function createPoliticalAnalysisAndPost(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY が設定されていません。');
  }

  const prompt = `
あなたは、政治ニュースをファクトベースで整理し、X投稿文案を作る編集者です。

目的:
左派・リベラル側の主張や行動について、事実・推測・疑問点を分けた上で、読者が「この主張や行動は本当に妥当なのか？」と考える投稿文を作る。

重要:
- 元情報に書かれていないことを事実として断定しない
- 不明な点は「確認が必要」「可能性がある」に留める
- 人格攻撃、属性攻撃、根拠のない犯罪断定はしない
- 批判対象は人物の属性ではなく、行動・制度・構図・利害関係にする
- 投稿文は保守寄り、疑問提起型、やや強め
- ハッシュタグは使わない
- 最後にURLを添える
- 出力は必ずJSONのみ

元記事URL:
${data.url}

タイトル:
${data.title}

元情報:
${data.sourceText}

次のJSON形式で出力してください。

{
  "fact": "確認できる事実を2〜4行で整理",
  "inference": "推測・見立て。断定せず、可能性として整理",
  "question": "読者に投げる疑問点",
  "interest": "利害関係や制度上の論点。不明なら「追加確認が必要」",
  "history": "過去行動や関連背景。不明なら「追加確認が必要」",
  "post": "X投稿文。200〜600字程度。最後にURLを入れる"
}
`;

  const payload = {
    model: 'gpt-4.1-mini',
    input: prompt,
    temperature: 0.7
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const text = response.getContentText();
  const dataJson = JSON.parse(text);

  if (response.getResponseCode() >= 400) {
    throw new Error(dataJson.error ? dataJson.error.message : text);
  }

  const outputText = dataJson.output_text || dataJson.output
    .flatMap(item => item.content || [])
    .map(content => content.text || '')
    .join('')
    .trim();

  return JSON.parse(outputText);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI投稿文')
    .addItem('ニュース候補を収集', 'collectNewsCandidates')
    .addItem('選択行の投稿文を作成', 'generateSelectedCandidatePosts')
    .addToUi();
}
function generateUnprocessedCandidatePosts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('投稿候補');
  if (!sheet) {
    throw new Error('投稿候補 シートが見つかりません。');
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = name => headers.indexOf(name) + 1;

  const aiCol = col('AI投稿文');
  const statusCol = col('ステータス');

  if (!aiCol || !statusCol) {
    throw new Error('AI投稿文 または ステータス 列が見つかりません。');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let row = 2; row <= lastRow; row++) {
    const aiPost = sheet.getRange(row, aiCol).getValue();
    const status = sheet.getRange(row, statusCol).getValue();

    if (aiPost) continue;
    if (status && status !== '候補') continue;

    const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
    generateCandidatePostsForRange(sheet, range);
  }
}
