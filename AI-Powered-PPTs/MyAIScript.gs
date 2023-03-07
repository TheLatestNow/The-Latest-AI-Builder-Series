// MODEL_TYPE which specifies the OpenAI model to use
// MY_API_KEY which retrieves the API key from the script properties
const MODEL_TYPE = "gpt-3.5-turbo";
const MY_API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');

// GoogleSlides Parametrers
const presentation = SlidesApp.getActivePresentation();
const slide = presentation.getSelection().getCurrentPage();
const ui = SlidesApp.getUi();

// CreateMenu function creates a button in the menubar of the Google Docs
function onOpen() {
  ui.createMenu("MyAI")
  .addItem("Write presentation outline", "callOpenAI")
  .addItem("Create Images", "generateImage")
  .addToUi();
}

function setRequestOptions(requestBody){
    const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + MY_API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };
  return requestOptions
}

function getSlideDimension(){
  // Get the slide dimensions
  const slideWidth = presentation.getPageWidth();
  const slideHeight = presentation.getPageHeight();

  // Calculate the dimensions and position of the footer shape
  const footerHeight = 200; // change this to the desired height of the footer
  const footerX = 10;
  const footerY = slideHeight - footerHeight;
  const footerWidth = slideWidth - 10;

  // Create the text box inside the footer shape
  const textBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, footerX, footerY, footerWidth, footerHeight);
  return textBox;
}

function callOpenAI(){
  const temperature = 0;
  const maxTokens = 4000;

  const presentation = SlidesApp.getActivePresentation();
  const selectedText = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
  const gptPrompt = "generate 5 bullet points outline for: " + selectedText;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{ role: "user", content: gptPrompt }],
    temperature,
    max_tokens: maxTokens,
  };

  var requestOptions = setRequestOptions(requestBody)

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);

  const textRange = getSlideDimension().getText();
  textRange.setText(generatedText.toString());

  // Set the properties of the text box
  textRange.getTextStyle().setFontSize(14); // change this to the desired font size of the text box
  textRange.getTextStyle().setBold(true); // change this to the desired font weight of the text box
  textRange.getTextStyle().setForegroundColor('#5865f2'); // change this to the desired font color of the text box
  // textRange.getTextStyle().setBackgroundColor('#e9f6ff'); // change this to the desired font bg color of the text box
}

function generateImage() {
  const imagePrompt = presentation.getSelection().getCurrentPage().getShapes()[0].getText().asString();
    
  const requestBody2 = {
    "prompt": imagePrompt,
    "n": 1,
    "size": "512x512"
  };

  var requestOptions = setRequestOptions(requestBody2)
  const response2 = UrlFetchApp.fetch("https://api.openai.com/v1/images/generations", requestOptions);

  // Parse the response and get the generated image
  var responseImage = response2.getContentText();
  var json = JSON.parse(responseImage);
  var imageUrl=json['data'][0]['url']
  var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  slide.insertImage(imageBlob);
}

///////////////////////////// Copyright Disclaimer /////////////////////////////
// The code provided by "TheLatestNow" is the intellectual property of "TheLatestNow.com" and is protected by international copyright laws. Allowed only for personal use. Duplication or distribution of this code for any possible commercial use, in whole or in part, without the express written consent of "TheLatestNow" is strictly prohibited. Any unauthorized use or reproduction of this code may result in legal action against the violator. This code is provided as is, without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement. "TheLatestNow" shall not be liable for any damages arising from the use or misuse of this code, including but not limited to direct, indirect, incidental, punitive, and consequential damages. Using this code, you agree to abide by these terms and conditions.
