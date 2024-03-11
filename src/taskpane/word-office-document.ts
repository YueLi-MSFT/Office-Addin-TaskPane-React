/* global Word console */

const insertWordText = async (text: string) => {
    // Write text to the document.
    try {
      await Word.run(async (context) => {
        let body = context.document.body;
        body.insertParagraph(text, Word.InsertLocation.end);
        await context.sync();
      });
    } catch (error) {
      console.log("Error: " + error);
    }
  };
  
  export default insertWordText;