const MODEL_TYPE = "gpt-3.5-turbo";
const MY_API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');

// createMenu function creates a button in the menubar of the Google Docs
function onOpen() {
  DocumentApp.getUi().createMenu("MyAI")
      .addItem("Write Blogpost Outline", "generateBlogpost")
      .addToUi();
}

function generateBlogpost() {
  const doc = DocumentApp.getActiveDocument();
  const userText = doc.getSelection().getRangeElements()[0].getElement().asText().getText();
  const body = doc.getBody();
  const prompt = "Write Blogpost Outline for " + userText;
  const temperature = 0;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
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
}
