const MODEL_TYPE = "gpt-3.5-turbo";
const MY_API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');

// GoogleDocs
const doc = DocumentApp.getActiveDocument();
// Define menu items and desired prompt
const items = [
  {label: 'Write Blog Outline', prompt: 'generate blog outline for:', functionName: 'menuItem0'},
  {label: 'Rephrase Sentence', prompt: 'rephrase the sentence:', functionName: 'menuItem1'},
  {label: 'Fix Grammer', prompt: 'correct grammer for:', functionName: 'menuItem2'},
  {label: 'Write Paragraph', prompt: 'write paragraph for:', functionName: 'menuItem3'},
  {label: 'Write YouTube Script', prompt: 'write youTube script with timestamp for:', functionName: 'menuItem4'},
  {label: 'Write SocialMedia Post', prompt: 'write social media post for:', functionName: 'menuItem5'},
  {label: 'Write Email Response', prompt: 'write email response:', functionName: 'menuItem6'}
];

const ui = DocumentApp.getUi();

function menuItem0() {generateResponse(items[0].label,items[0].prompt);}
function menuItem1() {generateResponse(items[1].label,items[1].prompt);}
function menuItem2() {generateResponse(items[2].label,items[2].prompt);}
function menuItem3() {generateResponse(items[3].label,items[3].prompt);}
function menuItem4() {generateResponse(items[4].label,items[4].prompt);}
function menuItem5() {generateResponse(items[5].label,items[5].prompt);}
function menuItem6() {generateResponse(items[6].label,items[6].prompt);}

// createMenu function creates a button in the menubar of the Google Docs
function onOpen() {
  const menu = ui.createMenu("MyAI-Pro");
  items.forEach(function(item) {
    menu.addItem(item.label, item.functionName);
  });
  menu.addSeparator;
  // menu.addItem('Show sidebar', 'showSidebar')
  menu.addToUi();
}

function generateResponse(selectedAction, selectedPrompt) {
  const selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText();
  
  let gptPrompt = selectedPrompt + " " + selectedText;
  Logger.log(gptPrompt);

  var alertResponse = ui.alert('AI is going to ' + gptPrompt, ui.ButtonSet.OK_CANCEL);
  
  if (alertResponse == ui.Button.OK) {
    callOpenAI(gptPrompt);
    } else {
      Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
      }
}

function callOpenAI(gptPrompt){
  const body = doc.getBody();
  const temperature = 0;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{ role: "user", content: gptPrompt }],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + MY_API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
  body.appendParagraph("...");
}

///////////////////////////// Copyright Disclaimer /////////////////////////////
// The code provided by "TheLatestNow" is the intellectual property of "TheLatestNow.com" and is protected by international copyright laws. Allowed only for personal use. Duplication or distribution of this code for any possible commercial use, in whole or in part, without the express written consent of "TheLatestNow" is strictly prohibited. Any unauthorized use or reproduction of this code may result in legal action against the violator. This code is provided as is, without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement. "TheLatestNow" shall not be liable for any damages arising from the use or misuse of this code, including but not limited to direct, indirect, incidental, punitive, and consequential damages. By using this code, you agree to abide by these terms and conditions.
