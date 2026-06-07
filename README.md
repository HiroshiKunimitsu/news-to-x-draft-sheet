# News to X Draft Sheet

News to X Draft Sheet is a Google Sheets and Google Apps Script tool that collects public news items and generates AI-assisted draft posts for X.

The project is designed for a review-first workflow. It helps users collect public information, organize it in a spreadsheet, and generate draft social media posts, but it does not automatically publish posts.

## Features

- Collect news candidates from Google News RSS
- Store article titles, URLs, and summaries in Google Sheets
- Generate structured analysis with AI
- Create draft posts for X
- Track review status in a spreadsheet
- Avoid duplicate URLs
- Keep humans in control before posting

## Use Cases

- News monitoring
- Public policy tracking
- Content research
- Social media draft preparation
- Editorial workflow support

## Spreadsheet Structure

Create two sheets.

### Keyword Settings

Sheet name:

キーワード設定

Columns:

キーワード / 対象 / 優先度

Example:

外国人政策 / ニュース / 高  
教育 中立性 / ニュース / 高  
沖縄 基地 / ニュース / 中

### Post Candidates

Sheet name:

投稿候補

Columns:

取得日時 / 情報源 / 元記事URL / タイトル / 元情報 / 事実 / 推測 / 疑問点 / 利害関係 / 過去行動 / AI投稿文 / ステータス

## Setup

1. Create a Google Spreadsheet.
2. Add the two sheets described above.
3. Open Extensions > Apps Script.
4. Paste the code from Code.gs.
5. Set your OpenAI API key in Script Properties.

Script property name:

OPENAI_API_KEY

## Workflow

1. Add keywords to the キーワード設定 sheet.
2. Run collectNewsCandidates.
3. News items are added to the 投稿候補 sheet.
4. Run generateSelectedCandidatePosts or generateUnprocessedCandidatePosts.
5. AI-generated analysis and draft posts are added to the sheet.
6. Review the draft manually before posting.

## Safety Notes

This project is intended to generate drafts, not automatically publish content.

Users should verify facts, review sources, and edit generated text before posting. The tool separates facts, assumptions, questions, and draft text to support responsible editorial review.

## License

MIT
