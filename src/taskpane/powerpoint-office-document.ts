/* global Office console */

const insertPowerPointText = async (text: string) => {
  // Write text to the selected slide.
  try {
    Office.context.document.setSelectedDataAsync(text, (asyncResult: Office.AsyncResult<void>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        throw asyncResult.error.message;
      }
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export default insertPowerPointText;