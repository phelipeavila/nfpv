/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  // Your code goes here

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// The global variable
var g = {};
var p = {};
var trib = {};
var tabelas = [];
    //{ "index":1 , "num_linhas":5 , "linha_ini":9 , "linha_fin":16 }];
var colunas = [
  {ini: "B", fin: "K"},
  {ini: "M", fin: "M"},
  {ini: "P", fin: "R"},
  {ini: "T", fin: "T"},
  {ini: "V", fin: "W"},
  {ini: "Y", fin: "AB"},
  {ini: "AE", fin: "AF"},
  {ini: "AH", fin: "AI"},
  {ini: "AK", fin: "AK"}
]

var listUF = ["AC",  "AL",  "AM",  "AP",  "BA",  "CE",  "DF",  "ES",  "GO",  "MA",  "MT",  "MS",  "MG", "PA",  "PB",  "PR",  "PE",  "PI",  "RN",  "RS",  "RJ",  "RO",  "RR",  "SC",  "SP",  "SE",  "TO"];

// the add-in command functions need to be available in global scope
g.action = action;
