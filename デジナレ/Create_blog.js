const apiKey = 'APIKEY';
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("作成");
const system_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gpt');
const refer_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('refer');
const apiUrl = "https://api.openai.com/v1/chat/completions";
const input_data = sheet.getRange(2,2,5).getValues();
const input_system = system_sheet.getRange(3,2,10).getValues();
const input_title = refer_sheet.getRange(38,3,3).getValues();
const input_summary = refer_sheet.getRange(38,4,3).getValues();

function set(){
  const refer_titles = input_title.join(" }\n\n - { ");
  const refer_summary = input_summary.join(" }\n\n - { ");
  titles = " - { " + refer_titles + " }";
  summary = " - { " + refer_summary + " }";
  sheet.getRange('B5').setValue(titles);
  sheet.getRange('B6').setValue(summary);
}

function set_data(messages){
  const System_setting = input_system[0][0];
  const Keyword = 'keyWord[' + input_data[0][0] + ']';
  const Target = 'target audience[' + input_data[1][0] + ']';
  const Purpose = 'purpose of blogging[' + input_data[2][0] + ']';
  const Refer_title = 'Reference titles [ /n' + input_data[3][0] + ' ]'
  const Refer_summary = 'References [ /n' + input_data[4][0] + ' ]'
  const draft_plan = 'draft_plan[' + sheet.getRange('D2').getValue + ']';

  messages.push({'role': 'system', 'content': System_setting});
  messages.push({'role': 'system', 'content': Keyword});
  messages.push({'role': 'system', 'content': Target});
  messages.push({'role': 'system', 'content': Purpose});
  messages.push({'role': 'system', 'content': Refer_title});
  messages.push({'role': 'system', 'content': Refer_summary});
  messages.push({'role': 'system', 'content': draft_plan});
}

function Create_titles() {
  let messages = [];
  
  set_data(messages)

  messages.push({'role': 'user', 'content': input_system[1][0]});

  //OpenAIのAPIリクエストに必要なヘッダー情報を設定
  const headers = {
    'Authorization':'Bearer '+ apiKey,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };
  //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
  const options = {
    'muteHttpExceptions' : true,
    'headers': headers, 
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-4o',
      'max_tokens' : 1024,
      'temperature' : 0.9,
      'messages': messages})
  };
  //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
  const title_response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText()); 
  let titles = title_response.choices[0].message.content;
  //過去の投稿履歴に実行結果をスプレッドシートに追加
  sheet.getRange(7,2).setValue(titles);
}

function Create_draft() {
  let messages = [];

  set_data(messages)

  messages.push({'role': 'user', 'content': input_system[2][0]});
  messages.push({'role': 'user', 'content': sheet.getRange(7,2).getValue()});

  //OpenAIのAPIリクエストに必要なヘッダー情報を設定
  const headers = {
    'Authorization':'Bearer '+ apiKey,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };
  //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
  const options = {
    'muteHttpExceptions' : true,
    'headers': headers, 
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-4o',
      'max_tokens' : 1024,
      'temperature' : 0.9,
      'messages': messages})
  };
  //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
  const draft_response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText()); 
  let Draft = draft_response.choices[0].message.content;
  //過去の投稿履歴に実行結果をスプレッドシートに追加
  sheet.getRange(8,2).setValue(Draft);
}

function Create_Chapter(){
  let messages = [];
  set_data(messages);
  messages.push({'role': 'assistant', 'content': sheet.getRange(8,2).getValue()});

  const prompt_chapter1 = input_system[4][0];
  const prompt_chapter2 = input_system[5][0];
  const prompt_chapter3 = input_system[6][0];
  const prompt_chapter4 = input_system[7][0];
  const prompt_chapter5 = input_system[8][0];

  const prompt = [
    {prompt_message: prompt_chapter1, outputCell: 'B10'},
    {prompt_message: prompt_chapter2, outputCell: 'B11'},
    {prompt_message: prompt_chapter3, outputCell: 'B12'},
    {prompt_message: prompt_chapter4, outputCell: 'B13'},
    {prompt_message: prompt_chapter5, outputCell: 'B14'},
  ];

  prompt.forEach((item) =>{
    messages.push({'role': 'user', 'content': item.prompt_message});
    //OpenAIのAPIリクエストに必要なヘッダー情報を設定
    const headers = {
      'Authorization':'Bearer '+ apiKey,
      'Content-type': 'application/json',
      'X-Slack-No-Retry': 1
    };
    //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
    const options = {
      'muteHttpExceptions' : true,
      'headers': headers, 
      'method': 'POST',
     'payload': JSON.stringify({
       'model': 'gpt-4o',
       'max_tokens' : 1024,
       'temperature' : 0.9,
       'messages': messages})
    };
    //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
    const response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText()); 
    let output = response.choices[0].message.content;
    //過去の投稿履歴に実行結果をスプレッドシートに追加
    sheet.getRange(item.outputCell).setValue(output);
  });
}

function Create_startEnd(){
  let messages = [];
  set_data(messages);
  let message_log = sheet.getRange(10,2,5).getValues();
  let blog = message_log.join("\n\n");
  
  messages.push({'role': 'assistant', 'content': blog});
  messages.push({'role': 'assistant', 'content': sheet.getRange(8,2).getValue()});

  const prompt_intro = input_system[3][0];
  const prompt_matome = input_system[9][0];

  const prompt = [
    {prompt_message: prompt_intro, outputCell: 'B9'},
    {prompt_message: prompt_matome, outputCell: 'B15'},
  ];

  prompt.forEach((item) =>{
    messages.push({'role': 'user', 'content': item.prompt_message});
    //OpenAIのAPIリクエストに必要なヘッダー情報を設定
    const headers = {
      'Authorization':'Bearer '+ apiKey,
      'Content-type': 'application/json',
      'X-Slack-No-Retry': 1
    };
    //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
    const options = {
      'muteHttpExceptions' : true,
      'headers': headers, 
      'method': 'POST',
     'payload': JSON.stringify({
       'model': 'gpt-4o',
       'max_tokens' : 1024,
       'temperature' : 0.9,
       'messages': messages})
    };
    //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
    const response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText()); 
    let output = response.choices[0].message.content;
    //過去の投稿履歴に実行結果をスプレッドシートに追加
    sheet.getRange(item.outputCell).setValue(output);
  });

}

function createDocumentFromCell() {
  
  // A1セルの内容を取得
  const messeage_log = sheet.getRange(9,2,7).getValues();
  const content = messeage_log.join("\n\n");
  
  // ドキュメントの作成
  const title = sheet.getRange('B7').getValue();
  const doc = DocumentApp.create(title);
  
  // ドキュメントに内容を書き込み
  const body = doc.getBody();
  body.appendParagraph(content);
  
  // 作成したドキュメントのURLを取得
  const docUrl = doc.getUrl();
  
  // スプレッドシートにドキュメントのURLを表示
  sheet.getRange("B1").setValue(docUrl);
  
  Logger.log("Document created: " + docUrl);
}

/*
function RewriteDraft() {

  let messages = [];
  const input = sheet.getRange("H3").getValue();
  const pre_Draft = sheet.getRange("B8").getValue();
  const request = input + ' # 改善した構成案のみ出力して';

  messages.push({'role': 'user', 'content': request});
  messages.push({'role': 'assistant', 'content': pre_Draft});

  //OpenAIのAPIリクエストに必要なヘッダー情報を設定
  const headers = {
    'Authorization':'Bearer '+ apiKey,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };
  //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
  const options = {
    'muteHttpExceptions' : true,
    'headers': headers, 
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-4o',
      'max_tokens' : 1024,
      'temperature' : 0.9,
      'messages': messages})
  };
  //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
  const draft_response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText()); 
  let imp_draft = draft_response.choices[0].message.content;
  //過去の投稿履歴に実行結果をスプレッドシートに追加
  sheet.getRange(8,2).setValue(imp_draft);
}*/
/*
function RewriteChapter() {

  let messages = [];
  set_data(messages);
  const input = sheet.getRange("H4").getValue();
  const num_chapter = sheet.getRange('G4').getValue();

  const request = input + '# Output only the improved text.';
  var out_cell;

  if(num_chapter == 0){
    const pre_intro = sheet.getRange('B9').getValue();
    messages.push({'role': 'assistant', 'content': pre_intro});
    out_cell = 'B9';
  }else if(num_chapter == 1){
    pre_1 = sheet.getRange('B10').getValue();
    messages.push({'role': 'assistant', 'content': pre_1});
    out_cell = 'B10';
  }else if(num_chapter == 2){
    const pre_2 = sheet.getRange('B11').getValue();
    messages.push({'role': 'assistant', 'content': pre_2});
    out_cell = 'B11';
  }else if(num_chapter == 3){
    const pre_3 = sheet.getRange('B12').getValue();
    messages.push({'role': 'assistant', 'content': pre_3});
    out_cell = 'B12';
  }else if(num_chapter == 4){
    const pre_4 = sheet.getRange('B13').getValue();
    messages.push({'role': 'assistant', 'content': pre_4});
    out_cell = 'B13';
  }else if(num_chapter == 5){
    const pre_5 = sheet.getRange('B14').getValue();
    messages.push({'role': 'assistant', 'content': pre_5});
    out_cell = 'B14';
  }else if(num_chapter == 6){
    const pre_end = sheet.getRange('B15').getValue();
    messages.push({'role': 'assistant', 'content': pre_end});
    out_cell = 'B15';
  }
  
  messages.push({'role': 'user', 'content': request});

  //OpenAIのAPIリクエストに必要なヘッダー情報を設定
  const headers = {
    'Authorization':'Bearer '+ apiKey,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };
  //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
  const options = {
    'muteHttpExceptions' : true,
    'headers': headers, 
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-4o',
      'max_tokens' : 1024,
      'temperature' : 0.9,
      'messages': messages})
  };
  //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
  const draft_response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText()); 
  let imp_draft = draft_response.choices[0].message.content;
  //過去の投稿履歴に実行結果をスプレッドシートに追加
  sheet.getRange(out_cell).setValue(imp_draft);
}*/

function deleteChatHistory(){
  //過去の投稿履歴データを削除
  sheet.getRange(1,2,15).clearContent();  
}

function main(){
  Create_titles();
  Create_draft();
  Create_Chapter();
  Create_startEnd();
  createDocumentFromCell();
}