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
var perfil = {};
var g = {};
var p = {};

var tabelas = [];
const SECRET = '027504';
var DEBUG = false;
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

var paraExcluir = {
  inicial:-1,
  final:-1
}

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


var param = {};
param = {
  versao: 0,
  hoje: 0,
  dataCambio: 0,
  dolarPTAX: 0,
  euroPTAX : 0,
  txImpHW: 2.2,
  txImpSW: 1.2,
  ufDest: "",
  ufOrig: "",
  destGoiania: false,
  zerarICMS: false,
  aplicarComissaoDirGov: false,
  aplicarComissaoDirPriv: false,
  aplicarComissaoVP: false,
  aplicarComissaoGC: false,
  aplicarComissaoExec: false,
  aplicarComissaoPrev: false,
  aplicarComissaoParc: false,
  txAdm: 0,
  svTerc: 0,
  projetoEstrategico: false,
  politicaAutom: true,
  comissaoDirGov: 0,
  comissaoDirPriv: 0,
  comissaoVP: 0,
  comissaoGC: 0,
  comissaoExec: 0,
  comissaoPrev: 0,
  comissaoParc: 0,
  margem: 0,
  tipoFatur: ""
}

var trib = {};
trib = {
  irpjHW: 0.018,
  irpjSW: 0.078,
  csllHW: 0.0108,
  csllSW: 0.0288,
  cppHW: 0,
  cppSW: 0,
  issGYN: 0.03,
  issOut: 0.08,
  pis: 0.0065,
  cofins: 0.03,
  fatDireto: 0.11,
  tabelaIcms: []
}

var comissoes = {};
comissoes = {
    "diretoria_governo" : {
        "equal_0"     : 0,
        "btw_0_10"    : 0,
        "equal_10"    : 0,
        "btw_10_15"   : 0,
        "equal_15"    : 0.005,
        "btw_15_20"   : 0.005,
        "equal_20"    : 0.005,
        "btw_20_25"   : 0.005,
        "equal_25"    : 0.005,
        "btw_25_30"   : 0.005,
        "equal_30"    : 0.005,
        "btw_30_35"   : 0.005,
        "equal_35"    : 0.005,
        "btw_35_40"   : 0.005,
        "equal_40"    : 0.005,
        "btw_40_45"   : 0.005,
        "equal_45"    : 0.005,
        "btw_45_50"   : 0.005,
        "equal_gr_50" : 0.005
    },
    "diretoria_privado" : {
       "equal_0"     : 0,
       "btw_0_10"    : 0,
       "equal_10"    : 0,
       "btw_10_15"   : 0,
       "equal_15"    : 0.01,
       "btw_15_20"   : 0.01,
       "equal_20"    : 0.01,
       "btw_20_25"   : 0.01,
       "equal_25"    : 0.01,
       "btw_25_30"   : 0.01,
       "equal_30"    : 0.01,
       "btw_30_35"   : 0.01,
       "equal_35"    : 0.01,
       "btw_35_40"   : 0.01,
       "equal_40"    : 0.01,
       "btw_40_45"   : 0.01,
       "equal_45"    : 0.01,
       "btw_45_50"   : 0.01,
       "equal_gr_50" : 0.01
    },
    "vp_comercial" : {
        "equal_0"     : 0,
        "btw_0_10"    : 0,
        "equal_10"    : 0,
        "btw_10_15"   : 0,
        "equal_15"    : 0.005,
        "btw_15_20"   : 0.005,
        "equal_20"    : 0.005,
        "btw_20_25"   : 0.005,
        "equal_25"    : 0.005,
        "btw_25_30"   : 0.005,
        "equal_30"    : 0.005,
        "btw_30_35"   : 0.005,
        "equal_35"    : 0.005,
        "btw_35_40"   : 0.005,
        "equal_40"    : 0.005,
        "btw_40_45"   : 0.005,
        "equal_45"    : 0.005,
        "btw_45_50"   : 0.005,
        "equal_gr_50" : 0.005
    },
    "gerente_canais" : {
        "equal_0"     : 0,
        "btw_0_10"    : 0,
        "equal_10"    : 0,
        "btw_10_15"   : 0,
        "equal_15"    : 0.002,
        "btw_15_20"   : 0.002,
        "equal_20"    : 0.002,
        "btw_20_25"   : 0.002,
        "equal_25"    : 0.002,
        "btw_25_30"   : 0.002,
        "equal_30"    : 0.002,
        "btw_30_35"   : 0.002,
        "equal_35"    : 0.002,
        "btw_35_40"   : 0.002,
        "equal_40"    : 0.002,
        "btw_40_45"   : 0.002,
        "equal_45"    : 0.002,
        "btw_45_50"   : 0.002,
        "equal_gr_50" : 0.002
    },
    "executivo" : {
        "equal_0"     : 0,
        "btw_0_10"    : 0,
        "equal_10"    : 0.01,
        "btw_10_15"   : 0.01,
        "equal_15"    : 0.01,
        "btw_15_20"   : 0.01,
        "equal_20"    : 0.01,
        "btw_20_25"   : 0.02,
        "equal_25"    : 0.02,
        "btw_25_30"   : 0.02,
        "equal_30"    : 0.03,
        "btw_30_35"   : 0.03,
        "equal_35"    : 0.03,
        "btw_35_40"   : 0.03,
        "equal_40"    : 0.03,
        "btw_40_45"   : 0.03,
        "equal_45"    : 0.03,
        "btw_45_50"   : 0.03,
        "equal_gr_50" : 0.03
    },
    "parceiro" : {
        "equal_0"     : 0,
        "btw_0_10"    : 0,
        "equal_10"    : 0.03,
        "btw_10_15"   : 0.03,
        "equal_15"    : 0.08,
        "btw_15_20"   : 0.08,
        "equal_20"    : 0.12,
        "btw_20_25"   : 0.12,
        "equal_25"    : 0.15,
        "btw_25_30"   : 0.15,
        "equal_30"    : 0.15,
        "btw_30_35"   : 0.15,
        "equal_35"    : 0.15,
        "btw_35_40"   : 0.15,
        "equal_40"    : 0.15,
        "btw_40_45"   : 0.15,
        "equal_45"    : 0.15,
        "btw_45_50"   : 0.15,
        "equal_gr_50" : 0.15
    },
    "prevendas" : {
        "projeto_comum"    : 0.0015,
        "projeto_estrategico"  : 0.0025
    }
}