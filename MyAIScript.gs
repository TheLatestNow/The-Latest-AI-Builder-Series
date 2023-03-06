// MODEL_TYPE which specifies the OpenAI model to use
// MY_API_KEY which retrieves the API key from the script properties
const MODEL_TYPE = "gpt-3.5-turbo";
const MY_API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');

// CreateMenu function creates a button in the menubar of the Google Docs
function onOpen() {
  DocumentApp.getUi().createMenu("MyAI")
      .addItem("Write Blogpost Outline", "generateBlogpost")
      .addToUi();
}

function generateBlogpost() {
  
  // Get the active Google Docs document and the user-selected text
  const doc = DocumentApp.getActiveDocument();
  const userText = doc.getSelection().getRangeElements()[0].getElement().asText().getText();
  
  // Get the body of the document and the prompt for the OpenAI API request
  const body = doc.getBody();
  const prompt = "Write Blogpost Outline for " + userText;
  
  // Set the parameters for the OpenAI API request
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

  // Call the OpenAI API to generate the blog post outline
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  
  // Append the generated blog post outline to the document body
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
}
