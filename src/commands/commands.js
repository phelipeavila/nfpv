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
  {ini: "B", qtde: "H", fin: "K"},
  {ini: "M", fin: "M"},
  {ini: "P", fin: "R"},
  {ini: "T", fin: "T"},
  {ini: "V", fin: "W"},
  {ini: "Y", fin: "AB"},
  {ini: "AE", fin: "AF"},
  {ini: "AH", fin: "AI"},
  {ini: "AK", fin: "AK"}
]

var id = {precificacao: "{5B74A0A4-C313-D74D-B6C9-894790A73C89}",
          servicos:["{A7441363-1A72-4ACD-854A-C140198E488F}"],
          cronograma:"{4360F843-007A-4860-8658-B6E2AA8612CD}",
          despesas:"{894AD57B-5D6A-4C25-BE36-2685E497F2CD}",
          custom:[],
          modelos:"{4AE48D1B-D021-694E-979A-1A4692BE21BF}",
          param:"{8CA1180C-FC57-B844-961A-30FBEFDE3919}",
          list:"{10ACBAF9-AE8C-5442-82A3-E6248D6BD589}",
          trib:"{AADBF709-2404-D044-B6C5-A0D2780F195E}",
          login:"{D00C8A5B-C396-D248-821C-23D397BE6BAE}"}

var listUF = ["AC",  "AL",  "AM",  "AP",  "BA",  "CE",  "DF",  "ES",  "GO",  "MA",  "MT",  "MS",  "MG", "PA",  "PB",  "PR",  "PE",  "PI",  "RN",  "RS",  "RJ",  "RO",  "RR",  "SC",  "SP",  "SE",  "TO"];

// the add-in command functions need to be available in global scope
g.action = action;
