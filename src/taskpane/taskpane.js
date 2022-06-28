/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = () => {

  p.state = {
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
    aplicarComDir: true,
    aplicarComCom: true,
    aplicarComPre: true,
    aplicarComPar: true,
    politicaParceiros: true,
    comDir: 0,
    comCom: 0,
    comPre: 0,
    comPar: 0,
    margem: 0,
    txAdm: 0,
    svTerc: 0,
    tipoFatur: ""
  }

  trib.state = {
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
    tabelaIcms: []
  }
  

    // Connect handlers

    document.getElementById("btn-login").onclick = onLogin;
    document.getElementById("sideload-msg").style.display = "none";
    
    document.getElementById("nav-button-edicao").onclick = showEdicao;
    document.getElementById("btn-add-linha").onclick = addLineButtonOnCLick;
    document.getElementById("btn-rem-linhas").onclick = remLineButtonOnCLick;
    document.getElementById("btn-tabela").onclick = tableButtonOnCLick;
    document.getElementById("btn-kit").onclick = kitButtonOnCLick;
    document.getElementById("btn-contrib").onclick = contribButtonOnCLick;
    document.getElementById("btn-move-esq").onclick = mvLeftButtonOnClick;
    document.getElementById("btn-move-dir").onclick = mvRightButtonOnClick;
    document.getElementById("btn-add-plan-sv").onclick = addSheetServOnClick;
    document.getElementById("btn-add-plan-br").onclick = addSheetBlankOnClick;
    document.getElementById("btn-rem-plan").onclick = remSheetOnCLick;
    document.getElementById("input-lista-planilhas").onkeyup = planilhaSV;
    

    document.getElementById("nav-button-cam-fat").onclick = showCamFat;
    document.getElementById('content-cambio-input-data').onchange = cambioInputOnChange;
    document.getElementById('content-cambio-input-data').onclick = cambioInputOnChange;
    document.getElementById('content-cambio-btn-atualiza').onclick = lastPtaxQuotation;
    document.getElementById('content-cambio-input-usd').onchange = manualCambioInputOnChange;
    document.getElementById('content-cambio-input-eur').onchange = manualCambioInputOnChange;
    document.getElementById('content-cambio-select-uf-orig').onchange = selectUFFaturamentoOnChange;
    document.getElementById('content-cambio-select-uf-dest').onchange = selectUFFaturamentoOnChange;
    document.getElementById('content-cambio-select-tipo-fat').onchange = selectUFFaturamentoOnChange;
    document.getElementById('content-cambio-check-icms').onchange = selectUFFaturamentoOnChange;
    document.getElementById('content-cambio-dest-gyn').onchange = selectUFFaturamentoOnChange;
    
    document.getElementById("nav-button-margem").onclick = showMargem;
    document.getElementById("content-margem-input-margem").onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-dir-gov').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-vp-comercial').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-dir-priv').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-ger-canais').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-exec').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-prev').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-comissao-parc').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-proj-estrategico').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-check-politica-automatica').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-dir-gov') .onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-vp-comercial').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-dir-priv').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-ger-canais').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-exec').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-prev').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-comissao-parc').onchange = margemComissoesOnChange;
    document.getElementById('content-margem-input-tx-admin').onchange = loadFromMargemToSheet;
    document.getElementById('content-margem-input-sv-terc').onchange = loadFromMargemToSheet;

    document.getElementById("nav-button-fechamento").onclick = showFechamento;
    document.getElementById("content-fechamento-btn-siecon").onclick = sieconButtonOnClick;
    document.getElementById("content-fechamento-btn-cronograma").onclick = toggleButtonCronograma;
    document.getElementById("content-fechamento-btn-fechamento").onclick = toggleButtonFechamento;

    document.getElementById("footer-button-parametros").onclick = showParametros;
    document.getElementById("content-parametros-button-tributos").onclick = toggleButtonTrib;
    document.getElementById("content-parametros-button-listas").onclick = toggleButtonList;
    document.getElementById("content-parametros-button-param").onclick = toggleButtonParam;
    document.getElementById("content-parametros-input-tx-hw").onchange = loadFromParametrosToSheet;
    document.getElementById("content-parametros-input-tx-sw").onchange = loadFromParametrosToSheet;


    //esse bloco de código permite que o login seja inserido pressionando ENTER
    document.querySelector("#input-passwd").addEventListener("keyup", event => {
      if(event.key !== "Enter") return; // Use `.key` instead.
      document.querySelector("#btn-login").click(); // Things you want to do.
      event.preventDefault(); // No need to `return false;`.
    });

    
    onStart()

    //comeceAqui();
    //exibePlanilhas()

    var app_body = document.getElementById("app-body");
  
  linha = null;

};


async function comeceAqui() {
    escondeCampos();
    //buscar parametros da planilha
    await atualizaParametros();
    //buscar tributos da planilha
    await atualizaTributos();
    //atualizar campos do cambio no frontend
    //atualizaDivPTAX();
    //atualizar campos de margem no frontend
    //await buscaMargem();
    carregaListaUF()

    await Excel.run(async (context) => {
      var workbook = context.workbook;
      var sheet = workbook.worksheets.getItem(id.precificacao)
      var isFPV = true;


      try {
        sheet.load("name")
        await context.sync()  
      } catch (error) {
        log("não é FPV")
        isFPV = false;
      }
      
      if (isFPV){
        workbook.load("protection/protected");
        await context.sync()  

        if (!workbook.protection.protected) {
          workbook.protection.protect(SECRET);
         }
      }

  
      return context.sync();
      
    });
    
    await mostraOsPermitidos();

    
    //document.getElementById("perm-all-other-fields").hidden = true;
    //atualiza array com informações das tabelas
    //await atualizaArrayTabelas();
}


//atualiza da planilha para o frontend
function buscaMargem(){
  //document.getElementById('inputMargem').value = p.state.margem * 100;
  //document.getElementById('inputTxAdm').value = p.state.txAdm * 100;
  //document.getElementById('inputSvTerc').value = p.state.svTerc * 100;
  document.getElementById('inputMargem').value        = (Math.round(p.state.margem*1000000)/10000);
  document.getElementById('inputTxAdm').value         = (Math.round(p.state.txAdm*1000000)/10000);
  document.getElementById('inputSvTerc').value        = (Math.round(p.state.svTerc*1000000)/10000);
  document.getElementById('checkDiretoria').checked   = p.state.aplicarComDir;
  document.getElementById('checkComercial').checked   = p.state.aplicarComCom;
  document.getElementById('checkPrevendas').checked   = p.state.aplicarComPre;
  document.getElementById('checkParceiro').checked    = p.state.aplicarComPar;
  document.getElementById('checkParceiro').disabled    = !p.state.politicaParceiros;
  
  
  document.getElementById('inputTipoFaturamento').value       = p.state.tipoFatur;
  document.getElementById('inputUFOrigem').value              = p.state.ufOrig;
  document.getElementById('inputUFDestino').value             = p.state.ufDest;
  document.getElementById('checkGoiania').checked             = p.state.destGoiania;
  document.getElementById('checkICMS').checked                = p.state.zerarICMS;
  document.getElementById('checkPoliticaParceiros').checked   = p.state.politicaParceiros;
  
  document.getElementById("inputComissaoDiretoria").value      = p.state.comDir * 100;
  document.getElementById("inputComissaoComercial").value      = p.state.comCom * 100;
  document.getElementById("inputComissaoPrevendas").value      = p.state.comPre * 100;
  document.getElementById("inputComissaoParceiro").value       = p.state.comPar * 100;

  
}

async function atualizaComissoesManualmente(){

  if (p.state.aplicarComDir){
    p.state.comDir = arred4(document.getElementById("inputComissaoDiretoria").value / 100);
  }  

  if (p.state.aplicarComCom){
    p.state.comCom = arred4(document.getElementById("inputComissaoComercial").value / 100);
  }  
  if (p.state.aplicarComPre){
    p.state.comPre = arred4(document.getElementById("inputComissaoPrevendas").value / 100);
  }  
  if (p.state.aplicarComPar){
    p.state.comPar = arred4(document.getElementById("inputComissaoParceiro").value  / 100);
  }  


  await Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem(id.param);
    var  range = ws.getRange("B3:B6").load("values");
    //await context.sync();

    let novosvalores = [[p.state.comDir],
                        [p.state.comCom],
                        [p.state.comPre],
                        [p.state.comPar]
                    ];
    range.values = novosvalores;
    await context.sync();

  });

  buscaMargem();

}


//atualizar do frontend para planilha
async function atualizaMargem(){
    
    //p.state.margem = document.getElementById('inputMargem').value / 100;
    //p.state.txAdm = document.getElementById('inputTxAdm').value / 100;
    p.state.tipoFatur = document.getElementById('inputTipoFaturamento').value;
    p.state.ufOrig = document.getElementById('inputUFOrigem').value;
    p.state.ufDest = document.getElementById('inputUFDestino').value;

    p.state.margem = ((Math.round(document.getElementById('inputMargem').value*10000))/1000000);
    p.state.txAdm = ((Math.round(document.getElementById('inputTxAdm').value*10000))/1000000);
    p.state.svTerc = ((Math.round(document.getElementById('inputSvTerc').value*10000))/1000000);
    p.state.aplicarComDir = document.getElementById('checkDiretoria').checked;
    p.state.aplicarComCom = document.getElementById('checkComercial').checked;
    p.state.aplicarComPre = document.getElementById('checkPrevendas').checked;
    p.state.aplicarComPar = document.getElementById('checkParceiro').checked;
    p.state.destGoiania = document.getElementById('checkGoiania').checked;
    p.state.zerarICMS = document.getElementById('checkICMS').checked;
    p.state.politicaParceiros = document.getElementById('checkPoliticaParceiros').checked

    
    
    //calcula comissao diretoria
    if (!p.state.aplicarComDir) {
        p.state.comDir = 0;
    }else{
        if(p.state.margem <= 0) { // SE MARGEM <= 0%
            p.state.comDir = 0;
        }else if (p.state.margem > 0 && p.state.margem <= 0.10){ // SE 0% < MARGEM <= 10%   ->  1%
            p.state.comDir = 0.01;
        }else if (p.state.margem > 0.1 && p.state.margem <= 0.20){ // SE 10% < MARGEM <= 20%   ->  1%
            p.state.comDir = 0.01;
        }else {
            p.state.comDir = 0.01; // SE MARGEM > 15%   ->  1 %
        }
    }


    //calcula comissao comercial
    if (!p.state.aplicarComCom) {
        p.state.comCom = 0;
    }else{
        if(p.state.margem < 0.1) {
            p.state.comCom = 0;
        }else if (p.state.margem >= 0.1 && p.state.margem <= 0.20){
            p.state.comCom = 0.01;
        }else if (p.state.margem > 0.20 && p.state.margem <= 0.30){
            p.state.comCom = 0.02;
        }else {
            p.state.comCom = 0.03;
        }
    }

    //calcula comissao prevendas
    if (!p.state.aplicarComPre) {
      p.state.comPre = 0;
    }else{
      if(p.state.margem < 0.1) {
          p.state.comPre = 0.0025;
      }else if (p.state.margem >= 0.1 && p.state.margem <= 0.20){
          p.state.comPre = 0.0025;
      }else if (p.state.margem > 0.20 && p.state.margem <= 0.30){
          p.state.comPre = 0.0025;
      }else {
          p.state.comPre = 0.0025;
      }
    }  

    //calcula comissao parceiro
    if (!p.state.aplicarComPar) {
      p.state.comPar = 0;
    }else{
        if (p.state.politicaParceiros){
          if(p.state.margem < 0.1) {
              p.state.comPar = 0;
          }else if (p.state.margem >= 0.1 && p.state.margem < 0.15){
              p.state.comPar = 0.03;
          }else if (p.state.margem >= 0.15 && p.state.margem < 0.20){
              p.state.comPar = 0.08;
          }else if (p.state.margem >= 0.20 && p.state.margem < 0.25){
              p.state.comPar = 0.12;
          }else {
              p.state.comPar = 0.15;
          }
        }   
    }
    //

    if (p.state.ufDest != 'GO') {
      document.getElementById('checkGoiania').disabled = true;
      document.getElementById('checkGoiania').checked  = false;
      p.state.destGoiania = false;
      //checkGoiania
    } else {
      document.getElementById('checkGoiania').disabled = false;
    }
    
    await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem(id.param);
        var  range = ws.getRange("B1:B16").load("values");
        //await context.sync();

        let novosvalores = [[p.state.tipoFatur],
                            [p.state.margem],
                            [p.state.comDir],
                            [p.state.comCom],
                            [p.state.comPre],
                            [p.state.comPar],
                            [p.state.txAdm],
                            [p.state.svTerc],
                            [p.state.aplicarComDir],
                            [p.state.aplicarComCom],
                            [p.state.aplicarComPre],
                            [p.state.aplicarComPar],
                            [p.state.ufOrig],
                            [p.state.ufDest],
                            [p.state.destGoiania],
                            [p.state.zerarICMS]
                        ];
        range.values = novosvalores;
        await context.sync();
        
        //LINHA B22 - aplicar política de comissionamento para parceiros
        range = ws.getRange("B22").load("values");
        novosvalores = [[p.state.politicaParceiros]];
        range.values = novosvalores;
        await context.sync();
    });
    buscaMargem();
}


async function atualizaTributos() {
  Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem(id.trib);
    var  range = ws.getRange("B1:B22").load("values");
    //await context.sync();

    let valores_tributos = [[""],                 //1
                            [trib.state.irpjHW],  //2
                            [trib.state.irpjSW],  //3
                            [""],                 //4
                            [""],                 //5
                            [trib.state.csllHW],  //6
                            [trib.state.csllSW],  //7
                            [""],                 //8
                            [""],                 //9
                            [0],                  //10
                            [0],                  //11
                            [""],                 //12
                            [""],                 //13
                            [trib.state.issGYN],  //14
                            [trib.state.issOut],  //15
                            [""],                 //16
                            [""],                 //17
                            [trib.state.pis],     //18
                            [""],                 //19
                            [""],                 //20
                            [""],                 //21
                            [trib.state.cofins]  //22
                          ];

    range.values = valores_tributos;
    //range.values[1][1] = trib.state.irpjHW;
    //range.values[2][1] = trib.state.irpjSW;
    //range.values[5][1] = trib.state.csllHW;
    //range.values[6][1] = trib.state.csllSW;
    //range.values[9][1] = trib.state.cppHW;
    //range.values[10][1] = trib.state.cppSW;
    //range.values[13][1] = trib.state.issGYN;
    //range.values[14][1] = trib.state.issOut;
    //range.values[17][1] = trib.state.pis;
    //range.values[21][1] = trib.state.cofins;

    trib.state.tabelaIcms = ws.getRange("F3:AF29").load("values");
    await context.sync();

    trib.state.tabelaIcms = trib.state.tabelaIcms.m_values;


    //                            destino        x      origem
    //trib.state.tabelaIcms[listUF.indexOf("PB")][listUF.indexOf("RJ")]

    //log( trib.state.tabelaIcms);

    return context;
  })
}


//Lê os dados da planilha 'param' e carrega para as variáveis globais


async function atualizaParametros() {
  Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem(id.param);
    var  range = ws.getRange("B1:B22").load("values");
    await context.sync();

    p.state.hoje = new Date();    
    
    p.state.tipoFatur         = range.m_values[0][0];
    p.state.margem            = range.m_values[1][0];
    p.state.comDir            = range.m_values[2][0];
    p.state.comCom            = range.m_values[3][0];
    p.state.comPre            = range.m_values[4][0];
    p.state.comPar            = range.m_values[5][0];
    p.state.txAdm             = range.m_values[6][0];
    p.state.svTerc            = range.m_values[7][0];
    p.state.aplicarComDir     = range.m_values[8][0];
    p.state.aplicarComCom     = range.m_values[9][0];
    p.state.aplicarComPre     = range.m_values[10][0];
    p.state.aplicarComPar     = range.m_values[11][0];
    p.state.ufOrig            = range.m_values[12][0];
    p.state.ufDest            = range.m_values[13][0];
    p.state.destGoiania       = range.m_values[14][0];
    p.state.zerarICMS         = range.m_values[15][0];
    p.state.txImpHW           = range.m_values[16][0];
    p.state.txImpSW           = range.m_values[17][0];
    p.state.dataCambio        = ExcelDateToJSDate(range.m_values[18], -3);
    p.state.dolarPTAX         = range.m_values[19];
    p.state.euroPTAX          = range.m_values[20];
    p.state.politicaParceiros = simNaotoBoolean(range.m_values[21][0], true);

    

    //ATUALIZA ARRAY ID.SERVICOS
    range = ws.getRange("V3:V12").load("values");
    await context.sync();

    var a = range.values;
        
    for (i in a){
        if (a[i] != ''){
            id.servicos.push(a[i][0]);
        }
    }

    //ATUALIZA ARRAY ID.CUSTOM
    range = ws.getRange("Y2:Y11").load("values");
    await context.sync();

    var b = range.values;
        
    for (i in b){
        if (b[i] != ''){
            id.custom.push(b[i][0]);
        }
    }
    

    atualizaDivPTAX();
    buscaMargem();
    await atualizaListaPlanilhas()


    return context;
  })
}

async function ocultaPlanilhas(ID) {
  Excel.run(async (context)=>{

    var workbook = context.workbook;
    workbook.load("protection/protected");

    await context.sync();
      
    if (workbook.protection.protected) {
      workbook.protection.unprotect(SECRET);
    }

    const ws = context.workbook.worksheets;
    ws.getItem(ID).visibility = Excel.SheetVisibility.veryHidden ;
    

    workbook.protection.protect(SECRET);

    await context.sync();

    return context;
  })
}

async function exibePlanilhas(ID) {


  Excel.run(async (context)=>{

    var workbook = context.workbook;
    workbook.load("protection/protected");

    await context.sync();
      
    if (workbook.protection.protected) {
      workbook.protection.unprotect(SECRET);
    }

    const ws = context.workbook.worksheets;
    ws.getItem(ID).visibility = Excel.SheetVisibility.visible ;
    workbook.protection.protect(SECRET);
    
    ws.getItem(ID).activate();

    await context.sync();

    return context;
  })

}

async function getAllLogin(type = "") {
  return Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem("login");
    var  range = ws.getRange("A1:B20").load("values");
    await context.sync();
    return range.values;
  });
}

//retorna tabela com as permisssões de cada perfil de acesso
async function getPermissions(perfil) {
  var permissoes = await Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem("login");
    var  range = ws.getRange("F2:K20").load("values");
    await context.sync();
    return range.values;
  });

  //remove linhas em branco
  while (permissoes[permissoes.length -1][0] === ""){
    permissoes.pop();
  }

  switch (perfil) {
    case 10:
      for(i in permissoes){
        permissoes[i] = permissoes[i][0]
      }
      break;
    case 20:
      for(i in permissoes){
        permissoes[i] = permissoes[i][1]
      }
      break;
    case 30:
      for(i in permissoes){
        permissoes[i] = permissoes[i][2]
      }
      break;
    case 40:
      for(i in permissoes){
        permissoes[i] = permissoes[i][3]
      }
      break;
    case 50:
      for(i in permissoes){
        permissoes[i] = permissoes[i][4]
      }   
      break;
    case 60:
      for(i in permissoes){
        permissoes[i] = permissoes[i][5]
      }   
      break;
    default:
      break;
  }
  log(perfil)
  return permissoes;
}


//retorna o código de permissão (10, 20, 30, 40 ou 50). Se der erro, retorna -1.
async function getProfile(inputLogin = document.getElementById('input-passwd').value) {
  // var base = await getAllLogin();
  // var users = [];
  // var permis = [];

  // for (i in base){
  //   users[i] = base[i][0];
  //   permis[i] = base[i][1];
  // }
  
  // let perm = permis[users.indexOf(login)];

  // if (typeof perm == 'undefined' || perm == ''){
  //   return -1;
  // }
  // log(login)
  // return perm;

  var get_login = (async function(){
    let obj = null;
    let url = "https://phelipeavila.github.io/nfpv/users.json";

    try {
        obj = await (await fetch(url)).json();
    } catch(e) {
        log('error');
        log(url);
    }
    return obj;
  });


  let login = await get_login(); //JSON com todos os logins e acessos
  
  if ( !login.acessos.hasOwnProperty(inputLogin)) return -1;

  return 10;

}

//esconde todos os campos, com exceção do login
function escondeCampos(){
  document.getElementById("perm-login").hidden = false;
  document.getElementById("perm-cambio").hidden = true;
  document.getElementById("perm-faturamento").hidden = true;
  document.getElementById("perm-margens-comissoes").hidden = true;
  document.getElementById("perm-formatar").hidden = true;
  document.getElementById("perm-implantacao").hidden = true;
  document.getElementById("nav-ul").hidden = true;
}


//QUANDO CLICAR NO BOTÃO DE LOGIN CRIA A NAVBAR
async function mostraOsPermitidos(){
  

  var inputLogin = document.getElementById('input-passwd').value; //obtem o input do passwd no HTML
  
  //pega o JSON com os logins e permissoes
    var get_login = (async function(){
      let obj = null;
      let url = "https://phelipeavila.github.io/nfpv/users.json";

      try {
          obj = await (await fetch(url)).json();
      } catch(e) {
          log('error');
          log(url);
      }
      return obj;
    });


    let login = await get_login(); //JSON com todos os logins e acessos
  log(inputLogin);
  log(login.acessos.hasOwnProperty(inputLogin));
  

  if ( !login.usuarios.hasOwnProperty(inputLogin)) return -1; //se o login digitado não está no JSON, retorna -1

  perfil = login.usuarios[inputLogin];

  var permissoes = login.acessos[perfil]

  
  //atualiza a nav-bar adicionando os elementos da nav-ul
  //<li id="li-edicao" class="nav-item"><a id="a-edicao" href="#">Edição</a></li>
  //<li id="li-cambio" class="nav-item"><a id="a-cambio" href="#" style="font-size: 0.8rem;">Câmbio e<br> Faturamento</a></li>
  //<li id="li-margens" class="nav-item"><a id="a-margens" href="#">Margens</a></li>

  let li = (function (nome) {

    let listItem = document.createElement('li');
    let link = document.createElement('a');

    link.id = 'a-' + nome;
    
    switch (nome) {
      case 'edicao':
        link.text = 'Edição'
        break;
      case 'margens':
        link.text = 'Margens'
        break;
      case 'cambio':
        link.textContent = 'Câmbio e' + '\n' + 'Faturamento'
        break;
      case 'posvenda':
        link.text = 'Pós Venda'
        break;
      case 'config':
        link.text = 'Config.'
        break;
        
      default:
        break;
    }

    listItem.id = 'li-'+nome;
    listItem.className = 'nav-item';
    listItem.appendChild(link);

    return listItem;

  })

  

  if (login.acessos[perfil]['nav-edicao']){
    document.getElementById('nav-ul').appendChild(li('edicao'));
    document.getElementById("a-edicao").onclick = mostraEdicao;
  }

  if (login.acessos[perfil]['nav-faturamento']){
    document.getElementById('nav-ul').appendChild(li('cambio'));
    document.getElementById("a-cambio").onclick = mostraCambio;
  }
  if (login.acessos[perfil]['nav-margens']){
    document.getElementById('nav-ul').appendChild(li('margens'));
    document.getElementById("a-margens").onclick = mostraMargens;
  }
  if (login.acessos[perfil]['nav-posvenda']){
    document.getElementById('nav-ul').appendChild(li('posvenda'));
    document.getElementById("a-posvenda").onclick = mostraPosvenda;
  }
  if (login.acessos[perfil]['nav-config']){
    document.getElementById('nav-ul').appendChild(li('config'));
    document.getElementById("a-config").onclick = mostraPosvenda;
  }


  document.getElementById("perm-all-other-fields").hidden = false;
  document.getElementById("nav-ul").hidden = false;
  document.getElementById("perm-login").hidden = true;

  return permissoes;

}

async function atualizaListaPlanilhas(){
  await Excel.run(async (context)=>{
    //percorre o array de IDs e pula id.servicos[0], pois é a planilha modelo
    for (i in id.servicos){
      if (i > 0){

        //pega nome da planilha de serviços
        let plan = context.workbook.worksheets.getItem(id.servicos[i])
        plan.load("name");
        await context.sync();

        addOptionToList(plan.name)
      }
    }

    for (i in id.custom){
      //pega nome da planilha customizada/em branco
      let plan = context.workbook.worksheets.getItem(id.custom[i])
      plan.load("name");
      await context.sync();
      
      //atualiza lista de planlilhas em branco
      addOptionToList(plan.name, true)
    }

  });
}

//monitora o input em texto da lista de planilhas e habilita ou desabilita os botões
function planilhaSV(){
  const input = document.getElementById('input-lista-planilhas').value;
  const datalist = document.getElementById('datalist-planilhas').options;
  //log(inputSV.value);

  if (input == "SV" || input == "param" || input == "list" || input == "login" || input == "trib" ||  input == "extenso" ||  input == "Precificação" || input == "CRONOGRAMA-COMPRAS" || input == "ANEXO IV" || input == "modelos" || input == "DESPESAS-INDIRETAS" ){
    document.getElementById('btn-add-plan-sv').disabled = true;
    document.getElementById('btn-add-plan-br').disabled = true;
    document.getElementById('btn-rem-plan').disabled = true;
    return
  }

  if (input == ''){
    document.getElementById('btn-add-plan-sv').disabled = false;
    document.getElementById('btn-add-plan-br').disabled = false;
    document.getElementById('btn-rem-plan').disabled = true;
  }


  if (existeNaLista(input)){
    document.getElementById('btn-add-plan-sv').disabled = true;
    document.getElementById('btn-add-plan-br').disabled = true;
    document.getElementById('btn-rem-plan').disabled = false;
  }else{
    document.getElementById('btn-add-plan-sv').disabled = false;
    document.getElementById('btn-add-plan-br').disabled = false;
    document.getElementById('btn-rem-plan').disabled = true;
  }

}

//procura se um valor existe na lista de planilhas SV ou BR
//retorna true se existe e false se não
function existeNaLista (valor){
  const datalist = document.getElementById('datalist-planilhas').options;

  if (valor == ''){
    return false;
  }

  for (i in datalist){
    if (datalist[i].value == valor){
      return true;
    }
  }
  return false;
}

//recebe um texto e adiciona à lista de planlilhas
//a função deve recber como argumento o valor a ser adicionado.
//Caso o segundo argumento seja falso ou não seja informado, o valor é considerado Planilha de Serviços
//caso seja verdadeiro, o valor é considerado Planilha Customizada
function addOptionToList (valor, customizada = false){
  var node = document.createElement('option');
  node.value = valor;
  if(customizada){
    node.text = 'Planilha Customizada'
  }else{
    node.text = 'Planilha de Serviços'
  }


  //verifica se o valor já existe no datalist. Se existir, ignora
  if ( !existeNaLista (valor)){
    document.getElementById('datalist-planilhas').appendChild(node)
  }else{
    log("já existe")
  }
}

//mostra nav-li "Câmbio e Faturamento"
async function mostraCambio(){


  //pega o JSON com os logins e permissoes
  var get_login = (async function(){
    let obj = null;
    let url = "https://phelipeavila.github.io/nfpv/users.json";

    try {
        obj = await (await fetch(url)).json();
    } catch(e) {
        log('error');
        log(url);
    }
    return obj;
  });


  let login = await get_login(); //JSON com todos os logins e acessos


  //oculta todas os campos
  //oculta login
  document.getElementById("perm-login").hidden = true;
  //oculta margens
  document.getElementById("perm-margens-comissoes").hidden = true;
  //oculta formatar
  document.getElementById("perm-formatar").hidden = true;
  //oculta implantacao
  document.getElementById("perm-implantacao").hidden = true;
  


  //exibe o campo de PTAX
  document.getElementById("perm-cambio").hidden = false;
  //exibe faturamento
  document.getElementById("perm-faturamento").hidden = false;


  //se o usuário tem permissão somente de leitura
  //desabilita os campos
  if(login.acessos[perfil]["faturamento-so-leitura"]){
    document.getElementById("radioCambioData").disabled = true;
    document.getElementById("inputDate").disabled = true;
    document.getElementById("radioCambioManual").disabled = true;
    document.getElementById("inputUSD").disabled = true;
    document.getElementById("inputEUR").disabled = true;
    document.getElementById("inputUFOrigem").disabled = true;
    document.getElementById("inputUFDestino").disabled = true;
    document.getElementById("inputTipoFaturamento").disabled = true;
    document.getElementById("checkICMS").disabled = true;
    document.getElementById("checkGoiania").disabled = true;
  }

  var menu = document.getElementById("nav-ul");
  for (i in menu.children){
    if (/^([0-9]+)$/.test(i)){
      menu.children[i].style.borderBottom = ""
    }
    
  }
  menu.children["li-cambio"].style.borderBottom = "3px solid";
  menu.children["li-cambio"].style.borderColor = "white";

}

function mostraEdicao(){
  //oculta todas os campos
  //oculta login
  document.getElementById("perm-login").hidden = true;
  //oculta faturamento
  document.getElementById("perm-faturamento").hidden = true;
  //oculta margens
  document.getElementById("perm-margens-comissoes").hidden = true;

  //oculta implantacao
  document.getElementById("perm-implantacao").hidden = true;
  //oculta o campo de PTAX
  document.getElementById("perm-cambio").hidden = true;
  //exibe formatar
  document.getElementById("perm-formatar").hidden = false;
  
  
  var menu = document.getElementById("nav-ul");
  for (i in menu.children){
    if (/^([0-9]+)$/.test(i)){
      menu.children[i].style.borderBottom = ""
    }
    
  }
  menu.children["li-edicao"].style.borderBottom = "3px solid";
  menu.children["li-edicao"].style.borderColor = "white";


}

function mostraMargens(){
  //oculta todas os campos
  //oculta login
  document.getElementById("perm-login").hidden = true;
  //oculta faturamento
  document.getElementById("perm-faturamento").hidden = true;
  //oculta implantacao
  document.getElementById("perm-implantacao").hidden = true;
  //oculta o campo de PTAX
  document.getElementById("perm-cambio").hidden = true;
  //exibe formatar
  document.getElementById("perm-formatar").hidden = true;
  //exibe margens
  document.getElementById("perm-margens-comissoes").hidden = false;


  var menu = document.getElementById("nav-ul");
  for (i in menu.children){
    if (/^([0-9]+)$/.test(i)){
      menu.children[i].style.borderBottom = ""
    }
    
  }
  menu.children["li-margens"].style.borderBottom = "3px solid";
  menu.children["li-margens"].style.borderColor = "white";

  var inputs_comissoes = document.getElementsByClassName("somenteDiretoria"); //campu input no HTML que mostra os valores das comissões
  for ( i in inputs_comissoes){
    if (perfil == "administrador"){
      inputs_comissoes[i].hidden = false;
    }
  }

}

function mostraPosvenda(){
  //oculta todas os campos
  //oculta login
  document.getElementById("perm-login").hidden = true;
  //oculta faturamento
  document.getElementById("perm-faturamento").hidden = true;
  //exibe implantacao
  document.getElementById("perm-implantacao").hidden = false;
  //oculta o campo de PTAX
  document.getElementById("perm-cambio").hidden = true;
  //oculta formatar
  document.getElementById("perm-formatar").hidden = true;
  //oculta margens
  document.getElementById("perm-margens-comissoes").hidden = true;


  var menu = document.getElementById("nav-ul");
  for (i in menu.children){
    if (/^([0-9]+)$/.test(i)){
      menu.children[i].style.borderBottom = ""
    }
    
  }
  menu.children["li-posvenda"].style.borderBottom = "3px solid";
  menu.children["li-posvenda"].style.borderColor = "white";

}

// function openDialog() {
//   if (dialog != null){
//     dialog.close();
//   }
//   Office.context.ui.displayDialogAsync(
//     'https://localhost:3000/src/dialogs/popup.html',
//     { height: 45, width: 55 },

//     function (result) {
//       dialog = result.value;
//       dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
//       log(result.value)
//     }
//   );
// }

// function processMessage(arg) {
//   document.getElementById("user-name").innerHTML = arg.message;
//   log(arg.message)
//   dialog.close();
// }
// var dialog = null;

async function selectUFFaturamentoOnChange(){
  var select_uf_orig            = document.getElementById('content-cambio-select-uf-orig')
  var select_uf_dest            = document.getElementById('content-cambio-select-uf-dest')
  var select_tipo_faturamento   = document.getElementById('content-cambio-select-tipo-fat')
  var check_incentivo_icms      = document.getElementById('content-cambio-check-icms')
  var select_destino_goiania    = document.getElementById('content-cambio-dest-gyn')

  if (select_uf_dest.value != 'GO'){ 
    select_destino_goiania.checked = false;
    select_destino_goiania.disabled = true;
  }else{
    select_destino_goiania.disabled = false;
  }
  
  loadFromCambioToSheet()
}

async function lastPtaxQuotation(){
  var data_cambio = document.getElementById('content-cambio-input-data');
  data_cambio.value = ontemStringHTML();
  await cambioInputOnChange()

}

async function cambioInputOnChange(){
  var data_cambio = document.getElementById('content-cambio-input-data');
  var input_usd = document.getElementById('content-cambio-input-usd');
  var input_eur = document.getElementById('content-cambio-input-eur');

  input_usd.value = await ptaxQuotation('USD', data_cambio.value);
  input_eur.value = await ptaxQuotation('EUR', data_cambio.value);

  loadFromCambioToSheet();
}

function manualCambioInputOnChange(){
  var data_cambio = document.getElementById('content-cambio-input-data');
  data_cambio.value = '';

  loadFromCambioToSheet();
}

function showContent (item){
  
  var span_edicao     = document.getElementById('span-edicao');
  var span_cam_fat    = document.getElementById('span-cam-fat');
  var span_margem     = document.getElementById('span-margem');
  var span_fechamento = document.getElementById('span-fechamento');
  var span_parametros = document.getElementById('span-parametros');
  
  hideAllContent();


  switch (item) {
    case 0:
      showAllowedContent('content-parametros-button-tributos', 'content-parametros-button-tributos');
      showAllowedContent('content-parametros-button-listas', 'content-parametros-button-listas');
      showAllowedContent('content-parametros-button-param', 'content-parametros-button-param');
      showAllowedContent('content-parametros-input-tx-hw', 'content-parametros-input-tx-hw');
      showAllowedContent('content-parametros-input-tx-sw', 'content-parametros-input-tx-sw');
      span_parametros.hidden = false;
      break;
    case 1:
      span_edicao.hidden = false;
      break;
    case 2:
      showAllowedContent('content-cambio-input-data', 'content-cambio');
      showAllowedContent('content-cambio-btn-atualiza', 'content-cambio');
      showAllowedContent('content-cambio-input-usd', 'content-cambio');
      showAllowedContent('content-cambio-input-eur', 'content-cambio');
      showAllowedContent('content-cambio-select-uf-orig', 'content-cambio');
      showAllowedContent('content-cambio-select-uf-dest', 'content-cambio');
      showAllowedContent('content-cambio-select-tipo-fat', 'content-cambio');
      showAllowedContent('content-cambio-check-icms', 'content-cambio');
      showAllowedContent('content-cambio-dest-gyn', 'content-cambio');

      span_cam_fat.hidden = false;
      break;
    case 3:
      showAllowedContent('content-margem-input-margem', 'content-margem-input-margem');
      showAllowedContent('content-margem-check-comissao-dir-gov', 'content-margem-check-comissoes');
      showAllowedContent('content-margem-check-comissao-vp-comercial', 'content-margem-check-comissoes');
      showAllowedContent('content-margem-check-comissao-dir-priv', 'content-margem-check-comissoes');
      showAllowedContent('content-margem-check-comissao-ger-canais', 'content-margem-check-comissoes');
      showAllowedContent('content-margem-check-comissao-exec', 'content-margem-check-comissoes');
      showAllowedContent('content-margem-check-comissao-prev', 'content-margem-check-comissoes');
      showAllowedContent('content-margem-check-comissao-parc', 'content-margem-check-comissoes');

      showAllowedContent('content-margem-input-comissao-dir-gov', 'content-margem-input-comissoes');
      showAllowedContent('content-margem-input-comissao-vp-comercial', 'content-margem-input-comissoes');
      showAllowedContent('content-margem-input-comissao-dir-priv', 'content-margem-input-comissoes');
      showAllowedContent('content-margem-input-comissao-ger-canais', 'content-margem-input-comissoes');
      showAllowedContent('content-margem-input-comissao-exec', 'content-margem-input-comissoes');
      showAllowedContent('content-margem-input-comissao-prev', 'content-margem-input-comissoes');
      showAllowedContent('content-margem-input-comissao-parc', 'content-margem-input-comissoes');

      showAllowedContent('content-percent-comissao-dir-gov', 'content-margem-input-comissoes');
      showAllowedContent('content-percent-comissao-vp-comercial', 'content-margem-input-comissoes');
      showAllowedContent('content-percent-comissao-dir-priv', 'content-margem-input-comissoes');
      showAllowedContent('content-percent-comissao-ger-canais', 'content-margem-input-comissoes');
      showAllowedContent('content-percent-comissao-exec', 'content-margem-input-comissoes');
      showAllowedContent('content-percent-comissao-prev', 'content-margem-input-comissoes');
      showAllowedContent('content-percent-comissao-parc', 'content-margem-input-comissoes');
      
      showAllowedContent('content-margem-input-tx-admin', 'content-margem-input-tx-admin');
      showAllowedContent('content-margem-input-sv-terc', 'content-margem-input-sv-terc');
      showAllowedContent('content-margem-percent-sv-terc', 'content-margem-input-sv-terc');
      showAllowedContent('content-margem-label-sv-terc', 'content-margem-input-sv-terc');
      showAllowedContent('content-margem-check-proj-estrategico', 'content-margem-check-proj-estrategico');
      showAllowedContent('content-margem-label-proj-estrategico', 'content-margem-check-proj-estrategico');
      showAllowedContent('content-margem-check-politica-automatica', 'content-margem-check-politica-automatica');
      showAllowedContent('content-margem-label-politica-automatica', 'content-margem-check-politica-automatica');
      

      // if (param.versao == 1){
      //   document.getElementById('content-margem-check-comissao-vp-comercial').disabled = true;
      //   document.getElementById('content-margem-check-comissao-dir-priv').disabled = true;
      //   document.getElementById('content-margem-check-comissao-ger-canais').disabled = true;
      //   document.getElementById('content-margem-input-comissao-vp-comercial').disabled = true;
      //   document.getElementById('content-margem-input-comissao-dir-priv').disabled = true;
      //   document.getElementById('content-margem-input-comissao-ger-canais').disabled = true;
      // }
      
      span_margem.hidden = false;
      break;
      case 4:
        
        showAllowedContent('content-fechamento-btn-siecon', 'content-fechamento-btn-siecon');
        showAllowedContent('content-fechamento-btn-cronograma', 'content-fechamento-btn-cronograma');
        showAllowedContent('content-fechamento-btn-fechamento', 'content-fechamento-btn-fechamento');

      span_fechamento.hidden = false;
      break;
    default:
      log('showContent() - invalid param')
      break;
  }
}

function showAllowedContent (idHTML, acesso) {
  var permit = {
    'r': function () {
      document.getElementById(idHTML).style.display = 'block';
      document.getElementById(idHTML).disabled = true;
      return 0;
    },
    'w': function () {
      document.getElementById(idHTML).style.display = 'block'
      document.getElementById(idHTML).disabled = false;
      return 0;
    },
    true: function () {
      document.getElementById(idHTML).style.display = 'block';
      return 0;
    },
    false: function () {
      document.getElementById(idHTML).style.display = 'none'
     return 0;
    }
  };
  return permit[perfil.acessos[acesso]]();
}

function showParametros(){
  showContent(0);

  var div = document.getElementsByClassName('footer-a2')[0];
  div.style.borderTop = '2px solid #FF6347';
}

function showEdicao(){
  showContent(1);
  
  var div = document.getElementsByClassName('nav-a1')[0];
  div.style.borderBottom = '2px solid #FFFFFF';
}

function showCamFat(){
  showContent(2);

  var div = document.getElementsByClassName('nav-a2')[0];
  div.style.borderBottom = '2px solid #FFFFFF';
}

function showMargem(){
  showContent(3);

  var div = document.getElementsByClassName('nav-a3')[0];
  div.style.borderBottom = '2px solid #FFFFFF';
}

function showFechamento(){
  showContent(4);

  var div = document.getElementsByClassName('nav-a4')[0];
  div.style.borderBottom = '2px solid #FFFFFF';
}

function checkFPV(){
  return true;
}

//GV -> frontend
async function writeOnFrontend(){
  log(`writeOnFrontend()`)
  //carrega aba cambio e faturamento
  writeCambioOnFrontend();

  //carrega margem e comissoes
  writeMargemOnFrontend();

  //carrega tabelas sv e customizadas
  writeSheetsNamesOnFrontend();

  //carrega parametros
  //tx importacao
  writeParametrosOnFrontend();

}

//frontend (tudo) ->  GV -> planilha
async function loadFromFrontendToSheet(){
  log('loadFromFrontendToSheet()');

  loadFromEdicaoToSheet();
  loadFromCambioToSheet();
  loadFromMargemToSheet();
  loadFromParametrosToSheet();
  
  return 0
}

//exibe o menu superior e inferior
function showMenuItens (){

  var item_edicao     = document.getElementsByClassName('nav-a1')[0];
  var item_cambio     = document.getElementsByClassName('nav-a2')[0];
  var item_margem     = document.getElementsByClassName('nav-a3')[0];
  var item_fechamento = document.getElementsByClassName('nav-a4')[0];
  var item_parametros = document.getElementsByClassName('footer-a2')[0];

  perfil.acessos['nav-edicao'] ? item_edicao.style.display = 'block': item_edicao.style.display = 'none';
  perfil.acessos['nav-cambio'] ? item_cambio.style.display = 'block': item_cambio.style.display = 'none';
  perfil.acessos['nav-margem'] ? item_margem.style.display = 'block': item_margem.style.display = 'none';
  perfil.acessos['nav-fechamento'] ? item_fechamento.style.display = 'block': item_fechamento.style.display = 'none';
  perfil.acessos['nav-parametros'] ? item_parametros.style.display = 'block': item_parametros.style.display = 'none';

}

//oculta o conteúdo de todas as abas
function hideAllContent (){
  var span_login      = document.getElementById('span-login');
  var span_edicao     = document.getElementById('span-edicao');
  var span_cam_fat    = document.getElementById('span-cam-fat');
  var span_margem     = document.getElementById('span-margem');
  var span_fechamento = document.getElementById('span-fechamento');
  var span_parametros = document.getElementById('span-parametros');

  span_login.hidden       = true;
  span_edicao.hidden      = true;
  span_cam_fat.hidden     = true;
  span_margem.hidden      = true;
  span_fechamento.hidden  = true;
  span_parametros.hidden  = true;
  
  //span da navbar para mudar a cor da borda de baixo do menu
  var nav_a1    = document.getElementsByClassName('nav-a1')[0];
  var nav_a2    = document.getElementsByClassName('nav-a2')[0];
  var nav_a3    = document.getElementsByClassName('nav-a3')[0];
  var nav_a4    = document.getElementsByClassName('nav-a4')[0];
  var footer_a2 = document.getElementsByClassName('footer-a2')[0];
  
  nav_a1.style.borderBottom = '';
  nav_a2.style.borderBottom = '';
  nav_a3.style.borderBottom = '';
  nav_a4.style.borderBottom = '';
  footer_a2.style.borderTop = '';
}

//planilha -> GV -> frontend
async function loadFromSheet(){
  //valida versão do layout param da FPV
  var versao = await getFromSheet(id.param, 'C1', 'values')
  param.versao = versao[0][0];
  log(`versão FPV: ${versao[0][0]}`)

  switch (param.versao) {
    //se for v1
    case '':
      log(`campo param!C1 vazio -> V1`)
      await loadFromParamSheetV1();
      await loadFromTribSheetV1();
      break;
      
    //se for v1
    case 1:
      log(`campo param!C1 = 1 -> V1`)
      await loadFromParamSheetV1();
      await loadFromTribSheetV1();
      break;
    
    //se for v2
    case 2:
      log(`campo param!C1 = 2 -> V2`)
      await loadFromParamSheetV2();
      await loadFromTribSheetV2();
      break;
    
    default:
      log(`campo param!C1 contém valor inválido`)
      return -1;
  }

  await writeOnFrontend();

  return 0;
} 

//ao clicar no botão de login
async function onLogin(){
  log('onLogin()');
  var input_login = document.getElementById('input-passwd').value;

  //busca usuarios e permissoes cadastradas
  var get_login = (async function(){
    let obj = null;
    //let url = "https://phelipeavila.github.io/nfpv/users.json";
    let url = "../../users.json";

    try {
        obj = await (await fetch(url)).json();
    } catch(e) {
        log('error');
        log(url);
    }
    return obj;
  });

  let login = await get_login(); //JSON com todos os usuarios e permissoes

  //exibe no log para debug
  log(input_login);
  log(login.acessos.hasOwnProperty(input_login));

  //valida senha digitada
  //se o login digitado não está no JSON, retorna -1
  if ( !login.usuarios.hasOwnProperty(input_login)) return -1;

  //salva o usuario e as permissoes de acesso em variável global
  perfil.usuario = login.usuarios[input_login];
  perfil.acessos = login.acessos[perfil.usuario];
  
  //carrega param e trib para variáveis globais
  //carrega variáveis globais para frontend
  await loadFromSheet()

  //habilita menus
  showMenuItens();
  hideAllContent();

}

//ao iniciar
function onStart(){
  var span_login = document.getElementById('span-login');

  //valida se a planilha é FPV
  if (!checkFPV()) {
    log(`O arquivo excel não é uma FPV válida`);
    return 0;
  }  

  //mostra a tela de login
  span_login.hidden = false;

  //tenta fazer o login automático
  onLogin()
}


//AS FUNÇÕES ABAIXO DEVEM SER REESCRITAS.

async function addLineButtonOnCLick(){
  const numLinhas = parseInt(document.getElementById("input-num-linha").value);
  await atualizaArrayTabelas();
  return await Excel.run(async (context) =>{

      //Esse trecho inicial valida se o local selecionado está na planilha 'Precificação'
      //e dentro de uma tabela de orçamento. Se algum desses requisitos não for 
      //cumprido, a função retornará 1
      var cell = context.workbook.getActiveCell();
      var a = cell.getCellProperties({address: true});
      await context.sync();

      var linhaSelecionada = parseInt( a.value[0][0]["address"].split('!')[1].replace(/[A-Z]/g, '') );
      log(linhaSelecionada);

      if (a.value[0][0]["address"].split('!')[0] != 'Precificação'){
          log('Sheet not allowed');
          return 1;
      }

      const ws = context.workbook.worksheets.getItem(id.precificacao);
      var index_tabela = await estaEmTabela();
      //se fora da tabela
      if (index_tabela == -1){
          log('Range not allowed');
          return 1;
      }

      ////se dentro da tabela
      //se nas linhas do cabeçalho -> insere no final
      if(linhaSelecionada == tabelas[index_tabela -1].linha_ini || linhaSelecionada == tabelas[index_tabela -1].linha_ini + 1){
          ws.getRange(tabelas[index_tabela -1].linha_fin.toString().concat(":"+ (tabelas[index_tabela -1].linha_fin + numLinhas -1 ))).insert(Excel.InsertShiftDirection.down);
          ws.getRange(tabelas[index_tabela -1].linha_fin.toString().concat(":"+ (tabelas[index_tabela -1].linha_fin + numLinhas -1 ))).copyFrom("modelos!4:4");

          for (let i = tabelas[index_tabela -1].linha_fin ; i< tabelas[index_tabela -1].linha_fin + numLinhas ; i++){
              let range = ws.getRange(`M${i}`);
              let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  
              conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
              conditionalFormat.custom.format.fill.color = "#FCE4D6";
          }
          tabelas[index_tabela - 1].linha_fin += numLinhas ;

          //ATUALIZA A FORMULA SUBTOTAL (OBS: AQUI DEVE SER ESCRITA A FORMULA COMO EXCEL EM INGLES!! NÃO USAR PONTO-E-VIRGULA E NOMES EM PT-BR)
          ws.getRange("K"+ tabelas[index_tabela - 1].linha_fin).formulas =
                     [["=SUBTOTAL(9,K" +( tabelas[index_tabela - 1].linha_ini + 2) + ":K" + (tabelas[index_tabela - 1].linha_fin - 1) + ")"]];

          log('se nas linhas do cabeçalho -> insere no final')
          await context.sync();
          await atualizaArrayTabelas();
          await renumerar();
          return context.sync()
      }

      //se kit
      for (i in tabelas[index_tabela -1].kit){
          let finalDoKit = tabelas[index_tabela -1].kit[i].linha + tabelas[index_tabela -1].kit[i].subitens ;
          //se cabeçalho do kit -> normal acima do kit
          if (linhaSelecionada == tabelas[index_tabela -1].kit[i].linha){
              
              ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).insert(Excel.InsertShiftDirection.down);
              ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).copyFrom("modelos!4:4");

              for (let i = linhaSelecionada ; i< linhaSelecionada + numLinhas ; i++){
                  let range = ws.getRange(`M${i}`);
                  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      
                  conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
                  conditionalFormat.custom.format.fill.color = "#FCE4D6";
              }

              tabelas[index_tabela - 1].linha_fin += numLinhas ;
              //ATUALIZA A FORMULA SUBTOTAL (OBS: AQUI DEVE SER ESCRITA A FORMULA COMO EXCEL EM INGLES!! NÃO USAR PONTO-E-VIRGULA E NOMES EM PT-BR)
              //ws.getRange("K"+ tabelas[index_tabela - 1].linha_fin).formulas =
              //        [["=SUBTOTAL(9,K" +( tabelas[index_tabela - 1].linha_ini + 2) + ":K" + (tabelas[index_tabela - 1].linha_fin - 1) + ")"]];

              log('se cabeçalho do kit -> normal acima do kit')
              await context.sync();
              await atualizaArrayTabelas();
              await renumerar();
              return context.sync()
          }

          //se primeira linha -> insere sublinha e corrige fórmulas do cabeçalho do kit
          if (linhaSelecionada == tabelas[index_tabela -1].kit[i].linha + 1){
              let range = "";
              log(`linha selecionada: ${linhaSelecionada}`)
              log(`linha selecionada: ${tabelas[index_tabela -1].kit[i].linha + 1}`)
              log(`final do kit: ${finalDoKit}`)
              ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).insert(Excel.InsertShiftDirection.down);
              ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).copyFrom("modelos!11:11");

              for (let i = linhaSelecionada ; i< linhaSelecionada + numLinhas ; i++){
                  let range = ws.getRange(`M${i}`);
                  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      
                  conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
                  conditionalFormat.custom.format.fill.color = "#FCE4D6";
              }

              tabelas[index_tabela - 1].linha_fin += numLinhas ;
              tabelas[index_tabela -1].kit[i].subitens += numLinhas;
              //ATUALIZA A FORMULA SUBTOTAL (OBS: AQUI DEVE SER ESCRITA A FORMULA COMO EXCEL EM INGLES!! NÃO USAR PONTO-E-VIRGULA E NOMES EM PT-BR)
              //ws.getRange("K"+ tabelas[index_tabela - 1].linha_fin).formulas =
              //        [["=SUBTOTAL(9,K" +( tabelas[index_tabela - 1].linha_ini + 2) + ":K" + (tabelas[index_tabela - 1].linha_fin - 1) + ")"]];


              log('se primeira linha -> insere dentro kit')
              await context.sync();

              //ATUALIZA AS FÓRMULAS DO CABEÇALHO DO KIT
              //valor unitário venda
              range = "J" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,K${tabelas[index_tabela -1].kit[i].linha + 1}:K${finalDoKit + numLinhas})/qtde,0)`]]

              //valor total venda
              range = "K" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,K${tabelas[index_tabela -1].kit[i].linha + 1}:K${finalDoKit + numLinhas})/qtde,0)*qtde`]]

              //contribuição


              //valor unitário custo
              range = "Q" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,R${tabelas[index_tabela -1].kit[i].linha + 1}:R${finalDoKit + numLinhas})/qtde,0)`]]

              //valor total custo
              range = "R" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,R${tabelas[index_tabela -1].kit[i].linha + 1}:R${finalDoKit + numLinhas})/qtde,0)*qtde`]]

              //custo unitário com desconto
              range = "AE" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,AF${tabelas[index_tabela -1].kit[i].linha + 1}:AF${finalDoKit + numLinhas})/qtde,0)`]]

              //custo total com desconto
              range = "AF" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,AF${tabelas[index_tabela -1].kit[i].linha + 1}:AF${finalDoKit + numLinhas})/qtde,0)*qtde`]]

              //custo unitário com desconto + importação
              range = "AH" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,AI${tabelas[index_tabela -1].kit[i].linha + 1}:AI${finalDoKit + numLinhas})/qtde,0)`]]

              //custo total com desconto + importação
              range = "AI" + tabelas[index_tabela -1].kit[i].linha;
              log(range)
              ws.getRange(range).formulas = [[`=iferror(SUBTOTAL(9,AI${tabelas[index_tabela -1].kit[i].linha + 1}:AI${finalDoKit + numLinhas})/qtde,0)*qtde`]]
              
              await atualizaArrayTabelas();
              await renumerar();
              await context.sync();

              return context.sync()
          }

          //se meio do kit -> insere sublinha
          if (linhaSelecionada > tabelas[index_tabela -1].kit[i].linha && linhaSelecionada <= finalDoKit ){
              log(`linha selecionada: ${linhaSelecionada}`)
              log(`final do kit: ${finalDoKit}`)
              ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).insert(Excel.InsertShiftDirection.down);
              ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).copyFrom("modelos!11:11");

              for (let i = linhaSelecionada ; i< linhaSelecionada + numLinhas ; i++){
                  let range = ws.getRange(`M${i}`);
                  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      
                  conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
                  conditionalFormat.custom.format.fill.color = "#FCE4D6";
              }

              tabelas[index_tabela - 1].linha_fin += numLinhas ;
              
              log('se dentro do kit -> insere dentro kit')
              await context.sync();
              await atualizaArrayTabelas();
              await renumerar();
              return context.sync()
          }            

      }



      
      ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).insert(Excel.InsertShiftDirection.down);
      ws.getRange(linhaSelecionada.toString().concat(":"+ (linhaSelecionada + numLinhas - 1))).copyFrom("modelos!4:4");

      for (let i = linhaSelecionada ; i< linhaSelecionada + numLinhas ; i++){
          let range = ws.getRange(`M${i}`);
          let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);

          conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
          conditionalFormat.custom.format.fill.color = "#FCE4D6";
      }
      
      //ws.getRange(tabelas[index_tabela -1].linha_fin.toString().concat(":"+ (tabelas[index_tabela -1].linha_fin + numLinhas -1 ))).insert(Excel.InsertShiftDirection.down);
      //ws.getRange(tabelas[index_tabela -1].linha_fin.toString().concat(":"+ (tabelas[index_tabela -1].linha_fin + numLinhas -1 ))).copyFrom("modelos!4:4");
      tabelas[index_tabela - 1].linha_fin += numLinhas ;

      //ATUALIZA A FORMULA SUBTOTAL (OBS: AQUI DEVE SER ESCRITA A FORMULA COMO EXCEL EM INGLES!! NÃO USAR PONTO-E-VIRGULA E NOMES EM PT-BR)
      ws.getRange("K"+ tabelas[index_tabela - 1].linha_fin).formulas =
                          [["=SUBTOTAL(9,K" +( tabelas[index_tabela - 1].linha_ini + 2) + ":K" + (tabelas[index_tabela - 1].linha_fin - 1) + ")"]];
       
       
      await context.sync();
      await atualizaArrayTabelas();
      await renumerar();
      return context.sync()
      
  })


}

async function remLineButtonOnCLick(){
  
  return await Excel.run(async (context) => {

    var selecao = await selectedRange();
    var workbook = context.workbook;

    if(selecao.planilha != 'Precificação'){
        log('Fora da planilha')
        return -1
    }
    
    log(`if(paraExcluir.inicial == selecao.inicial && paraExcluir.final == selecao.final)`)
    log(`if(${paraExcluir.inicial} == ${selecao.inicial} && ${paraExcluir.final} == ${selecao.final})`)
    log(`${(paraExcluir.inicial == selecao.inicial && paraExcluir.final == selecao.final)}`)
    
    
    if(paraExcluir.inicial == selecao.inicial && paraExcluir.final == selecao.final){
        workbook.worksheets.getItem(id.precificacao).getRange(selecao.inicial + ':' + selecao.final).delete("Up");
        paraExcluir.inicial = -1;
        paraExcluir.final = -1;
        await context.sync();
        await renumerar();
        await atualizaArrayTabelas();

        return 0
    }

    await atualizaArrayTabelas();

    //trata o range selecionado
    for (i in tabelas){
        //se selecao é a (primeira || segunda || terceira ) && (última || penúltima)
        if ( (selecao.inicial == tabelas[i].linha_ini || selecao.inicial == tabelas[i].linha_ini +1 || selecao.inicial == tabelas[i].linha_ini +2) && 
        (selecao.final == tabelas[i].linha_fin || selecao.final == tabelas[i].linha_fin -1)){
            log('excluir tabela toda');
            log(`novo range: ${tabelas[i].linha_ini}:${tabelas[i].linha_fin + 1}`);
            selecao.inicial = tabelas[i].linha_ini;
            selecao.final = tabelas[i].linha_fin + 1;

            workbook.worksheets.getItem(id.precificacao).getRange(selecao.inicial + ':' + selecao.final).select();
            paraExcluir.inicial = selecao.inicial;
            paraExcluir.final = selecao.final;

        
        //se selecao é entre (primeira) && (<= última)
        }else if ( (selecao.inicial == tabelas[i].linha_ini) && (selecao.final <= tabelas[i].linha_fin)){
            log('excluir tabela toda')
            log(`novo range: ${tabelas[i].linha_ini}:${tabelas[i].linha_fin + 1}`)
            selecao.inicial = tabelas[i].linha_ini;
            selecao.final = tabelas[i].linha_fin + 1;
            workbook.worksheets.getItem(id.precificacao).getRange(selecao.inicial + ':' + selecao.final).select();
            paraExcluir.inicial = selecao.inicial;
            paraExcluir.final = selecao.final;

        //se selecao é entre terceira e penúltima
        }else if(selecao.inicial >= tabelas[i].linha_ini +2 && selecao.final <= tabelas[i].linha_fin -1){
            log('meio da tabela')

            
            
            
            for (j in tabelas[i].kit){
            //se a seleção é o cabeçalho de um kit -> expandir para todo o kit
            if(selecao.final == selecao.inicial && selecao.final == tabelas[i].kit[j].linha){
                selecao.final = tabelas[i].kit[j].linha + tabelas[i].kit[j].subitens;
                paraExcluir.inicial = selecao.inicial;
                paraExcluir.final = selecao.final;
            }


            //se a linha final está dentro de um kit -> expandir seleção para o kit todo
            //somente se a linha inicial estiver fora do kit!!
                if (selecao.final >= tabelas[i].kit[j].linha && selecao.final < tabelas[i].kit[j].linha + tabelas[i].kit[j].subitens && selecao.inicial < tabelas[i].kit[j].linha){
                    selecao.final = tabelas[i].kit[j].linha + tabelas[i].kit[j].subitens;
                    paraExcluir.inicial = selecao.inicial;
                    paraExcluir.final = selecao.final;
                }
            }

            workbook.worksheets.getItem(id.precificacao).getRange(selecao.inicial + ':' + selecao.final).select();
            paraExcluir.inicial = selecao.inicial;
            paraExcluir.final = selecao.final;
        }
    }
});



///a seleção é somente a última linha (do subtotal)? Se sim, não excluir e retornar
///o range contém o cabeçalho de uma tabela? Se sim, expandir a seleção até o final 
///o range contém o cabeçalho de um kit? Se sim, a última linha do range é maior que a última linha do kit? Se não, expandir até o final do kit
///(validar se as fórmulas do cabeçalho do kit se ajustam sozinhas)
///o range contém a última linha de algum kit? Se sim, ajustar a formatação 
///seleciona a(s) linha(s) e exclui
//renumerar

}

async function tableButtonOnCLick(){
  await atualizaArrayTabelas()
      
  //log(tabelas.length);

  if (tabelas.length > 0){
      var linhaInicioNovaTabela =  tabelas[tabelas.length -1].linha_fin + 1;
  }else{
      var linhaInicioNovaTabela =  8;
  }
  //log(linhaInicioNovaTabela);
  let linStr = linhaInicioNovaTabela.toString().concat(":"+ (linhaInicioNovaTabela + 8));

  //return linStr;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(id.precificacao);

    sheet.getRange(linStr).insert(Excel.InsertShiftDirection.up);
    sheet.getRange(linStr).copyFrom("modelos!1:9");
    await context.sync();
  });
  await atualizaArrayTabelas();
  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(id.precificacao);

    for (let i = tabelas[tabelas.length-1].linha_ini+2 ; i< tabelas[tabelas.length-1].linha_fin ; i++){
        let range = ws.getRange(`M${i}`);
        let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);

        conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
        conditionalFormat.custom.format.fill.color = "#FCE4D6";
    }
  });
  
  await renumerar(); 
}

async function kitButtonOnCLick(){
  await atualizaArrayTabelas();
  var indexTabela = await estaEmTabela();

  if (indexTabela == -1) {
    log("fora de tabelaa");
    return -1
  }

  return await Excel.run(async (context) => {

    var range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();

    //return 0;
    
    //log(range.address)
    var kit = await selectedRange();


    //valida se a seleção é ou está contida em um kit.
    //se verdadeiro, irá desfazer o kit da seguinte forma:
    //seleciona todas as linhas do  kit
    //altera a formatação (cores e bordas)
    //exclui o cabeçalho
    //renumera
    //atualiza array de tabelas
    //return
    for (i in tabelas[indexTabela - 1].kit){
        if (kit.inicial >= tabelas[indexTabela - 1].kit[i].linha && kit.final <= tabelas[indexTabela - 1].kit[i].linha + tabelas[indexTabela - 1].kit[i].subitens){

            kit.inicial = tabelas[indexTabela - 1].kit[i].linha;
            kit.final = parseInt(tabelas[indexTabela - 1].kit[i].linha) + parseInt(tabelas[indexTabela - 1].kit[i].subitens);

            for (j in colunas){
                let selecao_nova = context.workbook.worksheets.getItem(id.precificacao).getRange(
                    colunas[j].ini + kit.inicial.toString() + ":" +
                    colunas[j].fin + kit.final.toString()  );

                if ( j == 0){
                    selecao_nova.ungroup(Excel.GroupOption.byRows);
                }

                if (j > 1 && j < 6){
                    selecao_nova.format.fill.color = "#E8E9EC";

                }else{
                    selecao_nova.format.fill.color = "#FFFFFF";

                }
                selecao_nova.format.borders.getItem('EdgeBottom').color = "#000000";
                selecao_nova.format.borders.getItem('EdgeRight').color = "#000000";
                selecao_nova.format.borders.getItem('EdgeLeft').color = "#000000";
                selecao_nova.format.borders.getItem('EdgeTop').color = "#000000";
                selecao_nova.format.borders.getItem('InsideHorizontal').color = "#000000";
                selecao_nova.format.borders.getItem('InsideVertical').color = "#000000";
            }     


            
            let cabecalho = context.workbook.worksheets.getItem(id.precificacao).getRange(tabelas[indexTabela - 1].kit[i].linha + ':' + tabelas[indexTabela - 1].kit[i].linha);
            cabecalho.delete("Up");
            await context.sync();


            await atualizaArrayTabelas();
            await renumerar();

            await context.sync();

            return
        }
    }



    //valida se a linha inicial está dentro de um kit e a linha final fora desse kit
    //se verdadeiro, retorna -1
    for (i in tabelas[indexTabela - 1].kit){
        if ((kit.inicial >= tabelas[indexTabela - 1].kit[i].linha ) &&
            (kit.inicial <= tabelas[indexTabela - 1].kit[i].linha + tabelas[indexTabela - 1].kit[i].subitens) &&
             kit.final > tabelas[indexTabela - 1].kit[i].linha + tabelas[indexTabela - 1].kit[i].subitens){
            log(11)
            return -1
        }
    }
    
    //valida se a linha final está dentro de um kit e a linha inicial fora desse kit
    //se verdadeiro, retorna -1
    for (i in tabelas[indexTabela - 1].kit){
        if (kit.inicial < tabelas[indexTabela - 1].kit[i].linha && //linha inicial acima do cabeçalho
            (kit.final <= tabelas[indexTabela - 1].kit[i].linha + tabelas[indexTabela - 1].kit[i].subitens) && //linha final acima do último subitem
            (kit.final >= tabelas[indexTabela - 1].kit[i].linha)){ //linha final abaido do cabeçalho
                log(22)
                return -1
        }
    }

    //valdar se o range está contido em alguma tabela:
    //1 - identificar se a primeira linha está em alguma tabela e qual o index dessa tabela
    //2 - identificar se a ultima linha do range esta na mesma tabela                    
    //a primeira linha do range do kit deve ser pelo menos a segunda linha da tabela (1.2) e no máximo a penúltima linha (pois a última é o subtotal)
    if ((tabelas[indexTabela - 1].linha_ini + 3) <= kit.inicial  
        && (tabelas[indexTabela - 1].linha_fin -1 ) >= kit.inicial
        && (tabelas[indexTabela - 1].linha_fin -1 ) >= kit.final ){
            //log("Seleção OK");
            //agrupa as células
            range.group(Excel.GroupOption.byRows);
            await context.sync();

            //seleciona demais células do kit e muda a formatação
            log(`-Ajustando a formatação das linhas dos subitens`)
            for (i in colunas){

                var subitens = context.workbook.worksheets.getItem("Precificação").getRange(
                    colunas[i].ini + kit.inicial.toString() + ":" +
                    colunas[i].fin + kit.final.toString()  );
                subitens.format.fill.color = "#D9D9D9";
                subitens.format.borders.getItem('EdgeBottom').color = "#000000";
                subitens.format.borders.getItem('EdgeRight').color = "#000000";
                subitens.format.borders.getItem('EdgeLeft').color = "#000000";
                subitens.format.borders.getItem('EdgeTop').color = "#F2F2F2";
                subitens.format.borders.getItem('InsideHorizontal').color = "#F2F2F2";
                subitens.format.borders.getItem('InsideVertical').color = "#000000";
            }              

            //FORMATAÇÃO DO HEADER
            log(`-Ajustando a formatação do Cabeçalho do kit`)
            for (i in colunas){
                if (i > 0){
                    var head = context.workbook.worksheets.getItem("Precificação").getRange(
                        colunas[i].ini + (kit.inicial - 1).toString() + ":" +
                        colunas[i].fin + (kit.inicial - 1).toString()  );
                    
                    head.format.font.color = "#203764";
                    head.format.font.bold = true;
                    head.format.borders.getItem('EdgeTop').color = "#000000";
                }   
            }

            //ARRUMA A borda superior da abaixo do kit---------------------------------
            log(`-Ajustando a borda superior abaixo do kit`)
            for (i in colunas){
                var subitens = context.workbook.worksheets.getItem("Precificação").getRange(
                    colunas[i].ini + (kit.final + 1).toString() + ":" +
                    colunas[i].fin + (kit.final + 1).toString()  );
                subitens.format.borders.getItem('EdgeTop').color = "#000000";
            }

            //formulas do cabeçalho do kit
            //valor de venda unitario
            log(`-Ajustando as fórmulas`)
            var celula = context.workbook.worksheets.getItem("Precificação").getRange(
               "J" + (kit.inicial - 1).toString() + ":" +
               "J" + (kit.inicial - 1 ).toString());
            celula.formulas = [[`=iferror(SUBTOTAL(9,K${kit.inicial}:K${kit.final})/qtde,0)`]];
            log(`${"J" + (kit.inicial - 1).toString() + ":" +"J" + (kit.inicial - 1 ).toString()}`)
            log(`=iferror(SUBTOTAL(9,K${kit.inicial}:K${kit.final})/qtde,0)`)
            //valor de venda total do item
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "K" + (kit.inicial - 1).toString() + ":" +
                "K" + (kit.inicial - 1 ).toString());
            celula.formulas = [[`=iferror(SUBTOTAL(9,K${kit.inicial}:K${kit.final})/qtde,0)*qtde`]];
            log(`${"K" + (kit.inicial - 1).toString() + ":" + "K" + (kit.inicial - 1 ).toString()}`)
            log(`=iferror(SUBTOTAL(9,K${kit.inicial}:K${kit.final})/qtde,0)*qtde`)
            //moeda
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "P" + (kit.inicial - 1).toString() + ":" +
                "P" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"P" + (kit.inicial - 1).toString() + ":" + "P" + (kit.inicial - 1 ).toString()}`)
            log("")
            //valor custo unitario
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "Q" + (kit.inicial - 1).toString() + ":" +
                "Q" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];//`=iferror(SUBTOTAL(9,R${kit.inicial}:R${kit.final})/qtde,0)`]];
            log(`${"Q" + (kit.inicial - 1).toString() + ":" + "Q" + (kit.inicial - 1 ).toString()}`)
            log(`=iferror(SUBTOTAL(9,R${kit.inicial}:R${kit.final})/qtde,0)`)
            //valor custo total
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "R" + (kit.inicial - 1).toString() + ":" +
                "R" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];//`=iferror((SUBTOTAL(9,R${kit.inicial}:R${kit.final})/qtde)*qtde,0)`]];
            log(`${"R" + (kit.inicial - 1).toString() + ":" + "R" + (kit.inicial - 1 ).toString()}`)
            log(`=iferror((SUBTOTAL(9,R${kit.inicial}:R${kit.final})/qtde)*qtde,0)`)
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "T" + (kit.inicial - 1).toString() + ":" +
                "T" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"T" + (kit.inicial - 1).toString() + ":" + "T" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "V" + (kit.inicial - 1).toString() + ":" +
                "V" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"V" + (kit.inicial - 1).toString() + ":" + "V" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "W" + (kit.inicial - 1).toString() + ":" +
                "W" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"W" + (kit.inicial - 1).toString() + ":" + "W" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "Y" + (kit.inicial - 1).toString() + ":" +
                "Y" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"Y" + (kit.inicial - 1).toString() + ":" + "Y" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "Z" + (kit.inicial - 1).toString() + ":" +
                "Z" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"Z" + (kit.inicial - 1).toString() + ":" + "Z" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AA" + (kit.inicial - 1).toString() + ":" +
                "AA" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"AA" + (kit.inicial - 1).toString() + ":" + "AA" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AB" + (kit.inicial - 1).toString() + ":" +
                "AB" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"AB" + (kit.inicial - 1).toString() + ":" + "AB" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AE" + (kit.inicial - 1).toString() + ":" +
                "AE" + (kit.inicial - 1 ).toString());
            celula.formulas = [[`=iferror(SUBTOTAL(9,AF${kit.inicial}:AF${kit.final})/qtde,0)`]];
            log(`${"AE" + (kit.inicial - 1).toString() + ":" + "AE" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AF" + (kit.inicial - 1).toString() + ":" +
                "AF" + (kit.inicial - 1 ).toString());
            celula.formulas = [[`=iferror((SUBTOTAL(9,AF${kit.inicial}:AF${kit.final})/qtde)*qtde,0)`]];
            log(`${"AF" + (kit.inicial - 1).toString() + ":" + "AF" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AH" + (kit.inicial - 1).toString() + ":" +
                "AH" + (kit.inicial - 1 ).toString());
            celula.formulas = [[`=iferror(SUBTOTAL(9,AI${kit.inicial}:AI${kit.final})/qtde,0)`]];
            log(`${"AH" + (kit.inicial - 1).toString() + ":" + "AH" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AI" + (kit.inicial - 1).toString() + ":" +
                "AI" + (kit.inicial - 1 ).toString());
            celula.formulas = [[`=iferror((SUBTOTAL(9,AI${kit.inicial}:AI${kit.final})/qtde)*qtde,0)`]];
            log(`${"AI" + (kit.inicial - 1).toString() + ":" + "AI" + (kit.inicial - 1 ).toString()}`)
            log("")
            celula = context.workbook.worksheets.getItem("Precificação").getRange(
                "AK" + (kit.inicial - 1).toString() + ":" +
                "AK" + (kit.inicial - 1 ).toString());
            celula.formulas = [[""]];
            log(`${"AK" + (kit.inicial - 1).toString() + ":" + "AK" + (kit.inicial - 1 ).toString()}`)
            log("")
            log("início context.sync()")
            await context.sync()
            log("fim context.sync()")
        
        log("FIM novoKit()")
        log("chamando renumerar()")
        await renumerar();
        return kit;
    }


    //sheet.getRange(linStr).insert(Excel.InsertShiftDirection.down);
    //sheet.getRange(linStr).copyFrom("modelos!1:9");
    log("fora de tabela")
    return kit;
  });

  log("dentro tabela");
}

async function contribButtonOnCLick(){
  await calculaContribuicao();
}

async function mvRightButtonOnClick(){
  await moveParaDireita();
}

async function mvLeftButtonOnClick(){
  await moveParaEsquerda();
}

async function addSheetServOnClick(){

  const COLUNA_ID_SV = 'V'

  if(id.servicos.length > 10){
      log(`Não é possível adicionar mais que 10 planilhas de serviço`)
      return -1
  }

  await Excel.run(async (context) => {
     
      var workbook = context.workbook;
      workbook.load("protection/protected");

      await context.sync();
      
      if (workbook.protection.protected) {
          workbook.protection.unprotect(SECRET);
          await context.sync();
      }
      

      var sampleSheet = workbook.worksheets.getItem(id.servicos[0]); 
      var precificacao = workbook.worksheets.getItem(id.precificacao); 
      sampleSheet.visibility = Excel.SheetVisibility.hidden;
      var copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, precificacao);
      sampleSheet.visibility = Excel.SheetVisibility.veryHidden;
      
      sampleSheet.load("name");
      copiedSheet.load("name");

      precificacao.load("position");
      sampleSheet.load("id");
      
      copiedSheet.visibility = Excel.SheetVisibility.visible;

      
      await context.sync();

      copiedSheet.position = precificacao.position + 1;
      copiedSheet.visibility = Excel.SheetVisibility.visible;

      if (document.getElementById('input-lista-planilhas').value == '' ){
          var i = 1;
          while (i != -1){
              try {
                  copiedSheet.name = "Serviços (" + i + ")";
                  await context.sync();
                  i = -1;
              } catch  {
                  i = i + 1;
              }
          }
      }else{
          copiedSheet.name = document.getElementById('input-lista-planilhas').value;
          document.getElementById('input-lista-planilhas').value = "";
      }

      copiedSheet.activate();
      workbook.protection.protect(SECRET)
      copiedSheet.load("id")
      await context.sync();
      id.servicos.push(copiedSheet.id);

      context.workbook.worksheets.getItem(id.param).getRange(COLUNA_ID_SV+ (id.servicos.length + 1) +":"+COLUNA_ID_SV+(id.servicos.length + 1)).values = copiedSheet.id;
  
      log("ID: " + copiedSheet.visibility );//+ "' was copied to '" + copiedSheet.name + "'");

      await writeSheetsNamesOnFrontend()
  });
}

async function addSheetBlankOnClick(){

  const COLUNA_ID_CUST = 'Y'

  if(id.custom.length > 9){
      log(`Não é possível adicionar mais que 10 planilhas customizadas`)
      return -1
  }

  await Excel.run(async (context) => {
     
      var workbook = context.workbook;
      workbook.load("protection/protected");

      await context.sync();
      
      if (workbook.protection.protected) {
          workbook.protection.unprotect(SECRET);
      }

      let novaPlanilha = workbook.worksheets.add('Planilha')
      //novaPlanilha.position = 0;
      //let sampleSheet = workbook.worksheets.getItem(id.servicos[0]); 
      //let precificacao = workbook.worksheets.getItem(id.precificacao); 
      //let copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, precificacao);
  
      novaPlanilha.load("name");
      //copiedSheet.load("name");

      novaPlanilha.load("position");
      novaPlanilha.load("id");
      
      novaPlanilha.visibility = Excel.SheetVisibility.visible;

      await context.sync();
      novaPlanilha.position = 0;
      //copiedSheet.visibility = Excel.SheetVisibility.visible;

      if (document.getElementById('input-lista-planilhas').value == '' ){
          var i = 1;
          while (i != -1){
              try {
                  novaPlanilha.name = "Planilha (" + i + ")";
                  await context.sync();
                  i = -1;
              } catch  {
                  i = i + 1;
              }
          }
      }else{
          novaPlanilha.name = document.getElementById('input-lista-planilhas').value;
          document.getElementById('input-lista-planilhas').value = "";
      }

      novaPlanilha.activate();
      workbook.protection.protect(SECRET)
      
      await context.sync();
      id.custom.push(novaPlanilha.id);

      workbook.worksheets.getItem(id.param).getRange(COLUNA_ID_CUST+ (id.custom.length + 2) +":"+COLUNA_ID_CUST+(id.custom.length + 2)).values = novaPlanilha.id;
  
      log("ID: " + novaPlanilha.visibility );//+ "' was copied to '" + copiedSheet.name + "'");

      await writeSheetsNamesOnFrontend()
  });
}

async function remSheetOnCLick(){
  var nome = document.getElementById('input-lista-planilhas').value;
  const COLUNA_ID_SV = 'V'
  const COLUNA_ID_CUSTOM = 'Y'
  
  return await Excel.run(async (context)=>{
      //verifica se a planilha existe
      const workbook = context.workbook;
      const plan = workbook.worksheets.getItem(nome)
      workbook.load("protection/protected");

      plan.load("id")
      try {
          await context.sync();    
      } catch (error) {
          log("planilha não existe")
          return -1
      }

      //verifica se a planilha é de serviços
      if (id.servicos.indexOf(plan.id) == -1 && id.custom.indexOf(plan.id) == -1){
          log("Não é uma planilha de serviços/customizada")

          return -1
      }else if(plan.id == "{A7441363-1A72-4ACD-854A-C140198E488F}"){
          log("Não é possível excluir a planilha modelo")
          return -1
      }

      //remove do frontend
      let lista = document.getElementById('datalist-planilhas').options;
      for (i in lista){
          if (lista[i].value == nome){
              lista[i].remove();
          }
      }


      //remove ID do array
      if (id.servicos.indexOf(plan.id) == -1){
          id.custom.splice(id.custom.indexOf(plan.id), 1);

          //remove todas as IDs das planilhas de serviço da aba param
          //escreve os IDs do array de ID, que agora está atualizado
          const param = workbook.worksheets.getItem(id.param);
          let range = param.getRange(COLUNA_ID_CUSTOM+"3"+":"+COLUNA_ID_CUSTOM+"12");
          range.load("values");
          await context.sync();
          for (i in range.values){
              param.getRange(COLUNA_ID_CUSTOM+(3+parseInt(i))).values = ''
              //range.values[i] = ['']
          }

          for (i in id.custom){
              param.getRange(COLUNA_ID_CUSTOM+(3+parseInt(i))).values = id.custom[i]
          }
          await context.sync();
      }else{
          id.servicos.splice(id.servicos.indexOf(plan.id), 1);

          //remove todas as IDs das planilhas de serviço da aba param
          //escreve os IDs do array de ID, que agora está atualizado
          const param = workbook.worksheets.getItem(id.param);
          let range = param.getRange(COLUNA_ID_SV+"2"+":"+COLUNA_ID_SV+"12");
          range.load("values");
          await context.sync();
          for (i in range.values){
              param.getRange(COLUNA_ID_SV+(2+parseInt(i))).values = ''
              //range.values[i] = ['']
          }

          for (i in id.servicos){
              param.getRange(COLUNA_ID_SV+(2+parseInt(i))).values = id.servicos[i]
          }
          await context.sync();
      }

      //remove todas as IDs das planilhas de serviço da aba param
      //escreve os IDs do array de ID, que agora está atualizado
      const param = workbook.worksheets.getItem(id.param);
      var range = param.getRange(COLUNA_ID_SV+"2"+":"+COLUNA_ID_SV+"12");
      range.load("values");
      await context.sync();
      for (i in range.values){
          param.getRange(COLUNA_ID_SV+(2+parseInt(i))).values = ''
          //range.values[i] = ['']
      }

      for (i in id.servicos){
          param.getRange(COLUNA_ID_SV+(2+parseInt(i))).values = id.servicos[i]
      }
      await context.sync();

      //remove a planilha
      if (workbook.protection.protected) {
          workbook.protection.unprotect(SECRET);
      }
      plan.delete()
      workbook.protection.protect(SECRET)
      document.getElementById('input-lista-planilhas').value = '';
      document.getElementById('btn-add-plan-sv').disabled = false;
      document.getElementById('btn-add-plan-br').disabled = false;
      document.getElementById('btn-rem-plan').disabled = true;
  });
}

async function toggleButtonTrib(){
  var status = await getFromSheet(id.trib, '', 'visibility');

  if (status == 'Visible'){
    ocultaPlanilhas(id.trib);
    return 0;
  }

  exibePlanilhas(id.trib)

}

async function toggleButtonList(){
  var status = await getFromSheet(id.list, '', 'visibility');

  if (status == 'Visible'){
    ocultaPlanilhas(id.list);
    return 0;
  }

  exibePlanilhas(id.list)
}

async function toggleButtonParam(){
  var status = await getFromSheet(id.param, '', 'visibility');

  if (status == 'Visible'){
    ocultaPlanilhas(id.param);
    return 0;
  }

  exibePlanilhas(id.param)

}

async function toggleButtonCronograma(){
  await cronograma();
}

async function toggleButtonFechamento(){
  await copiaTabelaParaDI();
}

async function toggleButtonSiecon(){
  var status = await getFromSheet(id.siecon, '', 'visibility');

  if (status == 'Visible'){
    ocultaPlanilhas(id.siecon);
    return 0;
  }

  exibePlanilhas(id.siecon)
}


async function sieconButtonOnClick(){
  await atualizaArrayTabelas();

  const PRIMEIRA_LINHA_DESTINO = 3;
  const COLUNA_VALORES_UNITARIOS_ORIGEM = 'AE'
  const COLUNA_VALORES_UNITARIOS_DESTINO = 'J'
  const LINHA_CABECALHO = 2;
  const COLUNA_CABECALHO = 'B'; //SOMENTE COLUNA MAIS À ESQUERDA

  var coluna_ini = colunas[0].ini;
  var coluna_fin = colunas[0].fin;
  var headKit = [];
  var numLinhasValidas = tabelas[tabelas.length - 1].linha_fin - tabelas[0].linha_ini - 2 //array com o número de linhas válidas
  
  return await Excel.run(async (context)=>{
      context.workbook.protection.unprotect(SECRET);
      const precificacao = context.workbook.worksheets.getItem(id.precificacao); //planilha Precificação
      const cronograma = context.workbook.worksheets.getItem(id.siecon);  //planilha CRONOGRAMA
      cronograma.load("visibility");
      await context.sync();

      var linhaCronograma = cronograma.getRange("1" + ":500");
      var arrayFormulaItem = [["", ""]];
      var range = "";

      //se a planilha já estiver criada, ao pressionar o botão ela será escondida
      if (cronograma.visibility == Excel.SheetVisibility.visible){
          cronograma.visibility = Excel.SheetVisibility.veryHidden;
          context.workbook.protection.protect(SECRET)
          return context.sync();
      }

      linhaCronograma.clear();
      cronograma.visibility = Excel.SheetVisibility.visible;
      await context.sync();


      //seleciona as linhas de B até K (coluna_fin) de todas as linhas com conteúdo
      var origem = precificacao.getRange(
        coluna_ini + (tabelas[0].linha_ini + 2) + ":" +coluna_fin + (tabelas[tabelas.length-1].linha_fin));
        
      var offset = (tabelas[0].linha_ini + 2) - PRIMEIRA_LINHA_DESTINO; //linha original - linha destino

      //copia para a planilha CRONOGRAMA, coluna B, linha PRIMEIRA_LINHA
      cronograma.getRange("B"+ PRIMEIRA_LINHA_DESTINO).copyFrom(origem);

      //copia os valores unitários da aba precificação (sem fórmulas, somente valores)

      var valores_origem = await getFromSheet(id.precificacao,
        COLUNA_VALORES_UNITARIOS_ORIGEM + (tabelas[0].linha_ini + 2) + ':' +  COLUNA_VALORES_UNITARIOS_ORIGEM + tabelas[tabelas.length - 1].linha_fin,
        'values');

      writeOnSheet(valores_origem,
         id.siecon,
         COLUNA_VALORES_UNITARIOS_DESTINO + (tabelas[0].linha_ini + 2 - offset) + ':' +  COLUNA_VALORES_UNITARIOS_DESTINO + (tabelas[tabelas.length - 1].linha_fin - offset),
         'values')

      await context.sync();


      
      //corrige as fórmulas dos cabeçalhos dos kits, se houver
      // 
      //verifica cabeçalhos de kits
      for (i in tabelas){
        for (j in tabelas[i].kit){
            headKit.push(tabelas[i].kit[j]);
        }
      }

    if (headKit.length > 0){

      for (i in headKit){
        //cabeçalho

        range = "J" + (parseInt(headKit[i].linha - offset)) + ":" + "K" + (parseInt(headKit[i].linha - offset));
        arrayFormulaItem =  [["", ""]];
        writeOnSheet(arrayFormulaItem, id.siecon, range, 'formulas')

      }
       await context.sync();
    }
      
              
      //remove as linhas em branco
      for (i in tabelas){
         cronograma.getRange((tabelas[tabelas.length - i -1].linha_fin - offset) + ":" + (tabelas[tabelas.length - i - 1].linha_fin - offset + 3)).delete(Excel.DeleteShiftDirection.up);
      }
      await context.sync();


      linhaCronograma = cronograma.getRange(COLUNA_CABECALHO + LINHA_CABECALHO + ':' + nextLetterInAlphabet(COLUNA_CABECALHO, 9) + LINHA_CABECALHO)
      linhaCronograma.format.fill.color = "#FF561C";
      linhaCronograma.format.borders.getItem('EdgeBottom').color = "#000000";
      linhaCronograma.format.borders.getItem('EdgeTop').color = "#000000";
      linhaCronograma.format.borders.getItem('EdgeRight').color = "#000000";
      linhaCronograma.format.borders.getItem('EdgeLeft').color = "#000000";
      linhaCronograma.format.borders.getItem('InsideVertical').color = "#000000";
      linhaCronograma.format.font.color = "#FFFFFF";
      linhaCronograma.format.font.bold = true;
      linhaCronograma.format.font.size = 9;
      linhaCronograma.format.rowHeight = 30;
      linhaCronograma.format.autoIndent = true;

      
      cronograma.getRange("D:D").insert(Excel.InsertShiftDirection.right);
      await context.sync();

      const header = [[ 'ITEM', 'NATUREZA\nVENDA', 'NATUREZA\nCOMPRA', 'DESCRIÇÃO', 'FABRICANTE', 'MODELO', 'NCM', 'QTDE', 'UN', 'VALOR UNITÁRIO (R$)', 'VALOR TOTAL (R$)']] 
      writeOnSheet(header, id.siecon, COLUNA_CABECALHO + LINHA_CABECALHO + ':' + nextLetterInAlphabet(COLUNA_CABECALHO, 10) + LINHA_CABECALHO, 'formulas')

      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, -1) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, -1));
      linhaCronograma.format.columnWidth = 7 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 0) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 0));
      linhaCronograma.format.columnWidth = 7 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 1) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 2));
      linhaCronograma.format.columnWidth = 7.5 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 3) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 3));
      linhaCronograma.format.columnWidth = 45 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 4) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 5));
      linhaCronograma.format.columnWidth = 15 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 6) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 6));
      linhaCronograma.format.columnWidth = 10 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 7) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 8));
      linhaCronograma.format.columnWidth = 7 * 6.54;
      linhaCronograma = cronograma.getRange(nextLetterInAlphabet(COLUNA_CABECALHO, 9) + ":" + nextLetterInAlphabet(COLUNA_CABECALHO, 10));
      linhaCronograma.format.columnWidth = 15 * 6.54;
      
      cronograma.activate();
      context.workbook.protection.protect(SECRET)
      await context.sync();
      //log("ID: " + cronograma.name );

  });

}


async function layoutTableResumo(sheet, coluna, linha){

  return await Excel.run(async (context)=>{
      const SHEET = sheet;
      const ws = context.workbook.worksheets.getItem(SHEET);
      const PRIMEIRA_LINHA = linha;
      const PRIMEIRA_COLUNA = coluna;
      const COLOR_BLACK = '#000000';
      const COLOR_WHITE = '#FFFFFF';
      const COLOR_DARK_GRAY = '#6F6F6E';
      const COLOR_LIGHT_GRAY = '#D9D9D9';
      const COLOR_DARK_ORANGE = '#FF561C';
      const COLOR_LIGHT_ORANGE = '#F4B084';

      const HEADER = [['RESUMO', 'VALOR', 'PERCENTUAL']]
      const FIRST_COLUMN = [
            ['VALOR DO FATURAMENTO'],
            ['CUSTOS DE AQUISIÇÃO EM REAIS'],
            ['CUSTOS DE IMPORTAÇÃO'],
            ['CUSTOS DIRETOS DE MÃO DE OBRA PRÓPRIA'],
            ['CUSTOS COM SUBCONTRATAÇÕES, LOCAÇÕES E DESPESAS DIVERSAS'],
            ['CUSTOS COM LOGÍSTICA PARA EQUIPE DE ACOMPANHAMENTO'],
            ['CUSTOS COM LOGÍSTICA PARA EQUIPE DE EXECUÇÃO'],
            ['CUSTOS COM FRETES'],
            ['COMISSÕES'],
            ['IMPOSTOS TOTAIS'],
            ['CRÉDITO ICMS'],
            ['IMPOSTOS (MENOS CRÉDITO DE ICMS)'],
            ['SUBTOTAL CUSTOS DIRETOS (ORÇAMENTO DE EXECUÇÃO)'],
            ['SERVIÇOS DE TERCEIROS'],
            ['TAXA ADMINISTRATIVA'],
            ['MARGEM LÍQUIDA']
      ];

      var range = '';
      range = PRIMEIRA_COLUNA + PRIMEIRA_LINHA + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, (HEADER[0].length - 1) ) + PRIMEIRA_LINHA;
      writeOnSheet(HEADER, SHEET, range)

      range = PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 1) + ':' + PRIMEIRA_COLUNA  + (PRIMEIRA_LINHA + 1 + FIRST_COLUMN.length - 1);
      writeOnSheet(FIRST_COLUMN, SHEET, range)

      //SEGUNDA COLUNA
      //FORMAT CURRENCY
      range = nextLetterInAlphabet(PRIMEIRA_COLUNA, 1) + (PRIMEIRA_LINHA + 1) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 1) + (PRIMEIRA_LINHA + 1 + FIRST_COLUMN.length - 1)
      var linhaSelecionada = ws.getRange(range);
      
      linhaSelecionada.style = "Currency";
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      
      //TERCEIRA COLUNA FORMAT PERCENT
      range = nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 1) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 1 + FIRST_COLUMN.length - 1)
      linhaSelecionada = ws.getRange(range);
      linhaSelecionada.style = "Percent";
      linhaSelecionada.numberFormat = "0.00%";
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      linhaSelecionada.format.horizontalAlignment = 'Center';
      
      //CABEÇALHO
      //FUNDO PRETO, FONTE BRANCA, CENTRALIZADO, MARGEM INTERNA VERTICAL PONTILHADA
      range = PRIMEIRA_COLUNA + PRIMEIRA_LINHA + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + PRIMEIRA_LINHA
      linhaSelecionada = ws.getRange(range);
      linhaSelecionada.format.fill.color = COLOR_BLACK;
      linhaSelecionada.format.font.color = COLOR_WHITE;
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      linhaSelecionada.format.horizontalAlignment = 'Center';
      linhaSelecionada.format.borders.getItem('InsideVertical').style = 'Dash';
      linhaSelecionada.format.borders.getItem('InsideVertical').weight = "Hairline";
      linhaSelecionada.format.borders.getItem('InsideVertical').color = COLOR_BLACK;

      
      //linhas em cinza escuro, fonte branca, margem superior e inferior pretas, margem interna vertical branca
      range = PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 1) + ',' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 1) + ',' +
        PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 13) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 13) + ',' + 
        PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 16) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 16);
      
      linhaSelecionada = ws.getRanges(range);
      linhaSelecionada.format.fill.color = COLOR_DARK_GRAY;
      linhaSelecionada.format.font.color = COLOR_WHITE;
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      linhaSelecionada.format.borders.getItem('EdgeBottom').color = COLOR_BLACK;
      linhaSelecionada.format.borders.getItem('EdgeTop').color = COLOR_BLACK;
      linhaSelecionada.format.borders.getItem('EdgeBottom').weight = "Hairline";
      linhaSelecionada.format.borders.getItem('EdgeTop').weight = "Hairline";
      

      //celulas laranja escuro, fonte branca, margem interna horizontal branca, 
      range = PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 2) + ':' + PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 9) + ',' +
        PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 12) + ',' + 
        PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 14) + ':' + PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 15);
    
      linhaSelecionada = ws.getRanges(range);
      linhaSelecionada.format.fill.color = COLOR_DARK_ORANGE;
      linhaSelecionada.format.font.color = COLOR_WHITE;
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      linhaSelecionada.format.borders.getItem('InsideHorizontal').style = 'Dash';
      linhaSelecionada.format.borders.getItem('InsideHorizontal').weight = "Hairline";
      linhaSelecionada.format.borders.getItem('InsideHorizontal').color = COLOR_BLACK;


      //celulas laranja claro, fonte branca, margem bottom branca, 
      range = PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 10) + ',' + PRIMEIRA_COLUNA + (PRIMEIRA_LINHA + 11);
    
      linhaSelecionada = ws.getRanges(range);
      linhaSelecionada.format.fill.color = COLOR_LIGHT_ORANGE;
      linhaSelecionada.format.font.color = COLOR_WHITE;
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      linhaSelecionada.format.borders.getItem('EdgeBottom').style = 'Dash';
      linhaSelecionada.format.borders.getItem('EdgeBottom').weight = "Hairline";
      linhaSelecionada.format.borders.getItem('EdgeBottom').color = COLOR_BLACK;

      //celulas cinza claro, fonte branca, margem interna branca, 
      range = nextLetterInAlphabet(PRIMEIRA_COLUNA, 1) + (PRIMEIRA_LINHA + 2) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 12) + ',' +
      nextLetterInAlphabet(PRIMEIRA_COLUNA, 1) + (PRIMEIRA_LINHA + 14) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 15);
    
      linhaSelecionada = ws.getRanges(range);
      linhaSelecionada.format.fill.color = COLOR_LIGHT_GRAY;
      linhaSelecionada.format.font.color = COLOR_BLACK;
      linhaSelecionada.format.font.bold = true;
      linhaSelecionada.format.font.size = 10;
      linhaSelecionada.format.borders.getItem('InsideHorizontal').style = 'Dash';
      linhaSelecionada.format.borders.getItem('InsideVertical').style = 'Dash';
      linhaSelecionada.format.borders.getItem('InsideHorizontal').weight = "Hairline";
      linhaSelecionada.format.borders.getItem('InsideVertical').weight = "Hairline";
      linhaSelecionada.format.borders.getItem('InsideHorizontal').color = COLOR_WHITE;
      linhaSelecionada.format.borders.getItem('InsideVertical').color = COLOR_WHITE;



      //margem mais escura
      range = (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 1) + ',' +
              (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA + 2) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 2) + ',' +
              (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA + 3) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 9) + ',' +
              (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA + 10) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 12) + ',' +
              (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA + 13) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 13) + ',' +
              (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA + 14) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 15) + ',' +
              (PRIMEIRA_COLUNA) + (PRIMEIRA_LINHA + 16) + ':' + nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 16);
    
      linhaSelecionada = ws.getRanges(range);
      linhaSelecionada.format.borders.getItem('EdgeTop').style = 'Continuous';
      linhaSelecionada.format.borders.getItem('EdgeRight').style = 'Continuous';
      linhaSelecionada.format.borders.getItem('EdgeLeft').style = 'Continuous';
      linhaSelecionada.format.borders.getItem('EdgeBottom').style = 'Continuous';

      linhaSelecionada.format.borders.getItem('EdgeTop').weight = 'Medium';
      linhaSelecionada.format.borders.getItem('EdgeRight').weight = 'Medium';
      linhaSelecionada.format.borders.getItem('EdgeLeft').weight = 'Medium';
      linhaSelecionada.format.borders.getItem('EdgeBottom').weight = 'Medium';

      linhaSelecionada.format.borders.getItem('EdgeTop').color = COLOR_BLACK;
      linhaSelecionada.format.borders.getItem('EdgeRight').color = COLOR_BLACK;
      linhaSelecionada.format.borders.getItem('EdgeLeft').color = COLOR_BLACK;
      linhaSelecionada.format.borders.getItem('EdgeBottom').color = COLOR_BLACK;

  });
  

}


async function contentTableResumo(coluna, linha){
  log("Início resumo()");
  await atualizaArrayTabelas();

  return await Excel.run(async (context)=>{
    const SHEET = id.despesas;
    const PRIMEIRA_LINHA = linha;
    const PRIMEIRA_COLUNA = coluna;
    const RANGE_DOS_VALORES = 'D3:E18';
    const precificacao = context.workbook.worksheets.getItem(id.precificacao); 
    const ws = context.workbook.worksheets.getItem(SHEET);
    var range = '';
    var formula =Array.from(Array(16), () => new Array(2));
    

    // VALOR DO FATURAMENTO
    //formula[0][0] = '';
    formula[0][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 1}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`; 
    // CUSTO DE AQUISIÇÃO EM REAIS
    formula[1][0] = `=-subtotal(9, Precificação!${colunas[6].fin}${tabelas[0].linha_ini}:${colunas[6].fin}${tabelas[tabelas.length - 1].linha_fin})- SUM(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 4}:${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 8})`;

    formula[1][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 2}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`; 
    //CUSTOS DE IMPORTAÇÃO
    formula[2][0] = `=-(subtotal(9, Precificação!${colunas[7].fin}${tabelas[0].linha_ini}:${colunas[7].fin}${tabelas[tabelas.length - 1].linha_fin}) -subtotal(9, Precificação!${colunas[6].fin}${tabelas[0].linha_ini}:${colunas[6].fin}${tabelas[tabelas.length - 1].linha_fin}))`
    
    formula[2][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 3}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`;

    //CUSTOS DE SERVIÇOS
    if(id.servicos.length > 1){
      // CUSTOS DIRETOS DE MÃO DE OBRA PRÓPRIA  (3)
      for (i in id.servicos){
          log(i)
          let servico = context.workbook.worksheets.getItem(id.servicos[i]);
          servico.load("name");
          await context.sync();

          log(servico.name)

          if (i == 1){
              formula[3][0] = "=-('" + servico.name + "'" + "!SUBTOTAL_MAO_DE_OBRA_PROPRIA";
          }
          if (i > 1){
              formula[3][0] = formula[3][0] + " + '" + servico.name + "'" + "!SUBTOTAL_MAO_DE_OBRA_PROPRIA";
          }
          if (i == id.servicos.length-1){
              formula[3][0] = formula[3][0] + " )"

          }
      }
      log(formula[3][0])


      // CUSTOS COM SUBCONTRATACOES E DESPESAS DIVERSAS (4)
      for (i in id.servicos){
          log(i)
          let servico = context.workbook.worksheets.getItem(id.servicos[i]);
          servico.load("name");
          await context.sync();

          log(servico.name)

          if (i == 1){
              formula[4][0] = "=-('" + servico.name + "'" + "!SUBTOTAL_SUBCONTRATACOES_E_DESPESAS_DIVERSAS";
          }
          if (i > 1){
              formula[4][0] = formula[4][0] + " + '" + servico.name + "'" + "!SUBTOTAL_SUBCONTRATACOES_E_DESPESAS_DIVERSAS";
          }
          if (i == id.servicos.length-1){
              formula[4][0] = formula[4][0] + " )"

          }
      }


      // CUSTOS COM LOGÍSTICA PARA EQUIPE DE ACOMPANHAMENTO (5)
      for (i in id.servicos){
          log(i)
          let servico = context.workbook.worksheets.getItem(id.servicos[i]);
          servico.load("name");
          await context.sync();

          log(servico.name)

          if (i == 1){
              formula[5][0] = "=-('" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_ACOMPANHAMENTO";
          }
          if (i > 1){
              formula[5][0] = formula[5][0] + " + '" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_ACOMPANHAMENTO";
          }
          if (i == id.servicos.length-1){
              formula[5][0] = formula[5][0] + " )"

          }
      }


      // CUSTOS COM LOGÍSTICA PARA EQUIPE DE EXECUÇÃO (6)
      for (i in id.servicos){
          log(i)
          let servico = context.workbook.worksheets.getItem(id.servicos[i]);
          servico.load("name");
          await context.sync();

          log(servico.name)

          if (i == 1){
              formula[6][0] = "=-('" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_CAMPO";
          }
          if (i > 1){
              formula[6][0] = formula[6][0] + " + '" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_CAMPO";
          }
          if (i == id.servicos.length-1){
              formula[6][0] = formula[6][0] + " )"
          }
      }


      // CUSTOS COM FRETES (7)
      for (i in id.servicos){
          log(i)
          let servico = context.workbook.worksheets.getItem(id.servicos[i]);
          servico.load("name");
          await context.sync();

          log(servico.name)

          if (i == 1){
              formula[7][0] = "=-('" + servico.name + "'" + "!SUBTOTAL_FRETES";
          }
          if (i > 1){
              formula[7][0] = formula[7][0] + " + '" + servico.name + "'" + "!SUBTOTAL_FRETES";
          }
          if (i == id.servicos.length-1){
              formula[7][0] = formula[7][0] + " )"
          }
      }
      
    }else{
      formula[3][0] = "0";
      formula[4][0] = "0";
      formula[5][0] = "0";
      formula[6][0] = "0";
      formula[7][0] = "0";
    }

    formula[3][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 4}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`
    formula[4][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 5}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`
    formula[5][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 6}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`
    formula[6][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 7}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`
    formula[7][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 8}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`

    //COMISSOES
    formula[8][0] = `=(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 2)}${PRIMEIRA_LINHA + 9}*$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1})`
    formula[8][1] = `=-${param.comissaoDirGov + param.comissaoDirPriv + param.comissaoVP + param.comissaoGC + param.comissaoExec + param.comissaoPrev + param.comissaoParc}`;
    
    //IMPOSTOS TOTAIS (-)
    formula[9][0] = ''
    formula[9][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 10}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`

    //CRÉDITO ICMS
    var creditoEmReais = 0;
    var arrayCustos = precificacao.getRange(`${colunas[7].fin}${tabelas[0].linha_ini}:${colunas[7].fin}${tabelas[tabelas.length - 1].linha_fin}`);
    var arrayCredICMS = precificacao.getRange(`${colunas[5].fin}${tabelas[0].linha_ini}:${colunas[5].fin}${tabelas[tabelas.length - 1].linha_fin}`);
    arrayCredICMS.load("values");
    arrayCustos.load("values");
    await context.sync();


    ////remove os valores dos cabeçalhos dos kits
    var offset = tabelas[0].linha_ini
    for (i in tabelas){
        for (j in tabelas[i].kit){
            arrayCustos.values[tabelas[i].kit[j].linha - offset] = ['']
        }
    }

    ////varre os arrays multiplicando os valores
    for (i in arrayCustos.values){
        if (arred4(arrayCustos.values[i] * arrayCredICMS.values[i]) >= 0){
            creditoEmReais = creditoEmReais + arred4(arrayCustos.values[i] * arrayCredICMS.values[i]);
        }
    }
    formula[10][0] = `${arred4(creditoEmReais)}`;
    formula[10][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 11}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`

    //IMPOSTOS (MENOS CRÉDITO ICMS)
    formula[11][0] = `=${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1) + (PRIMEIRA_LINHA + 10)} + ${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1) + (PRIMEIRA_LINHA + 11)}`
    formula[11][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 12}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`

    //SUBTOTAL CUSTOS DIRETOS (ORÇAMENTO DE EXECUÇÃO)
    //formula[12][0] = ''
    formula[12][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 13}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`

    //SERVIÇOS DE TERCEIROS
    formula[13][0] = `=${nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 14)} * $${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${(PRIMEIRA_LINHA + 1)}`
    formula[13][1] = `=-${param.svTerc}`

    //TAXA ADMINISTRATIVA
    formula[14][0] = `=${nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 15)} * $${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${(PRIMEIRA_LINHA + 1)}`
    formula[14][1] = `=-${param.txAdm}`


    //MARGEM LÍQUIDA
    formula[15][0] = `=${nextLetterInAlphabet(PRIMEIRA_COLUNA, 2) + (PRIMEIRA_LINHA + 14)} * $${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${(PRIMEIRA_LINHA + 1)}`
    formula[15][1] = `=IFERROR(${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}${PRIMEIRA_LINHA + 16}/$${nextLetterInAlphabet(PRIMEIRA_COLUNA, 1)}$${PRIMEIRA_LINHA + 1},0)`

    writeOnSheet(formula, id.despesas, RANGE_DOS_VALORES, 'formulas');
    return context.sync()

  });

}


async function resumo(){
  log("Início resumo()");
  await atualizaArrayTabelas();
  const colunaresumo = "D"
  return await Excel.run(async (context)=>{
      const precificacao = context.workbook.worksheets.getItem(id.precificacao); 
      const despesas = context.workbook.worksheets.getItem(id.despesas);
      var range = "";

      // VALOR DO FATURAMENTO
      // CUSTOS DE AQUISIÇÃO EM REAIS
      despesas.getRange("D4").formulas = `=-subtotal(9, Precificação!${colunas[6].fin}${tabelas[0].linha_ini}:${colunas[6].fin}${tabelas[tabelas.length - 1].linha_fin}) - SUM(D6:D10)`;
      // CUSTOS DE IMPORTAÇÃO
      despesas.getRange("D5").formulas = `=-(subtotal(9, Precificação!${colunas[7].fin}${tabelas[0].linha_ini}:${colunas[7].fin}${tabelas[tabelas.length - 1].linha_fin}) - subtotal(9, Precificação!${colunas[6].fin}${tabelas[0].linha_ini}:${colunas[6].fin}${tabelas[tabelas.length - 1].linha_fin}))`;

      if(id.servicos.length > 1){
          // CUSTOS DIRETOS DE MÃO DE OBRA PRÓPRIA  
          for (i in id.servicos){
              log(i)
              let servico = context.workbook.worksheets.getItem(id.servicos[i]);
              servico.load("name");
              await context.sync();

              log(servico.name)

              if (i == 1){
                  range = "=-('" + servico.name + "'" + "!SUBTOTAL_MAO_DE_OBRA_PROPRIA";
                  //log(`range1: ${range}`)
              }
              if (i > 1){
                  range = range + " + '" + servico.name + "'" + "!SUBTOTAL_MAO_DE_OBRA_PROPRIA";
                  //log(`range1+: ${range}`)
              }
              if (i == id.servicos.length-1){
                  range = range + " )"

              }
          }
          log(range)
          despesas.getRange("D6").formulas = range;

          // CUSTOS COM SUBCONTRATAÇÕES, LOCAÇÕES E DESPESAS DIVERSAS
          for (i in id.servicos){
              log(i)
              let servico = context.workbook.worksheets.getItem(id.servicos[i]);
              servico.load("name");
              await context.sync();

              log(servico.name)

              if (i == 1){
                  range = "=-('" + servico.name + "'" + "!SUBTOTAL_SUBCONTRATACOES_E_DESPESAS_DIVERSAS";
                  //log(`range1: ${range}`)
              }
              if (i > 1){
                  range = range + " + '" + servico.name + "'" + "!SUBTOTAL_SUBCONTRATACOES_E_DESPESAS_DIVERSAS";
                  //log(`range1+: ${range}`)
              }
              if (i == id.servicos.length-1){
                  range = range + " )"

              }
          }
          despesas.getRange("D7").formulas = range;

          // CUSTOS COM LOGÍSTICA PARA EQUIPE DE ACOMPANHAMENTO 
          for (i in id.servicos){
              log(i)
              let servico = context.workbook.worksheets.getItem(id.servicos[i]);
              servico.load("name");
              await context.sync();

              log(servico.name)

              if (i == 1){
                  range = "=-('" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_ACOMPANHAMENTO";
                  //log(`range1: ${range}`)
              }
              if (i > 1){
                  range = range + " + '" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_ACOMPANHAMENTO";
                  //log(`range1+: ${range}`)
              }
              if (i == id.servicos.length-1){
                  range = range + " )"

              }
          }
          despesas.getRange("D8").formulas = range;

          // CUSTOS COM LOGÍSTICA PARA EQUIPE DE EXECUÇÃO 
          for (i in id.servicos){
              log(i)
              let servico = context.workbook.worksheets.getItem(id.servicos[i]);
              servico.load("name");
              await context.sync();

              log(servico.name)

              if (i == 1){
                  range = "=-('" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_CAMPO";
                  //log(`range1: ${range}`)
              }
              if (i > 1){
                  range = range + " + '" + servico.name + "'" + "!SUBTOTAL_LOGISTICA_COM_EQUIPE_DE_CAMPO";
                  //log(`range1+: ${range}`)
              }
              if (i == id.servicos.length-1){
                  range = range + " )"

              }
          }
          despesas.getRange("D9").formulas = range;

          // CUSTOS COM FRETES 
          for (i in id.servicos){
              log(i)
              let servico = context.workbook.worksheets.getItem(id.servicos[i]);
              servico.load("name");
              await context.sync();

              log(servico.name)

              if (i == 1){
                  range = "=-('" + servico.name + "'" + "!SUBTOTAL_FRETES";
                  //log(`range1: ${range}`)
              }
              if (i > 1){
                  range = range + " + '" + servico.name + "'" + "!SUBTOTAL_FRETES";
                  //log(`range1+: ${range}`)
              }
              if (i == id.servicos.length-1){
                  range = range + " )"

              }
          }
          despesas.getRange("D10").formulas = range;
      }else{
          range = 0
          despesas.getRange("D6").formulas = range;
          despesas.getRange("D7").formulas = range;
          despesas.getRange("D8").formulas = range;
          despesas.getRange("D9").formulas = range;
          despesas.getRange("D10").formulas = range;
      }
      // COMISSÕES
      despesas.getRange("E11").formulas = `=-${p.state.comCom + p.state.comDir + p.state.comPar + p.state.comPre}`;
      // IMPOSTOS
      //despesas.getRange("D12").formulas = ``;
      // SERVIÇOS DE TERCEIROS
      despesas.getRange("E13").formulas = `=-${p.state.svTerc}`;
      // TAXA ADMINISTRATIVA
      despesas.getRange("E14").formulas = `=-${p.state.txAdm}`;
      
      // CRÉDITO ICMS
      var creditoEmReais = 0;
      var arrayCustos = precificacao.getRange(`${colunas[7].fin}${tabelas[0].linha_ini}:${colunas[7].fin}${tabelas[tabelas.length - 1].linha_fin}`);
      var arrayCredICMS = precificacao.getRange(`${colunas[5].fin}${tabelas[0].linha_ini}:${colunas[5].fin}${tabelas[tabelas.length - 1].linha_fin}`);
      arrayCredICMS.load("values");
      arrayCustos.load("values");
      await context.sync();


      ////remove os valores dos cabeçalhos dos kits
      var offset = tabelas[0].linha_ini
      for (i in tabelas){
          for (j in tabelas[i].kit){
              arrayCustos.values[tabelas[i].kit[j].linha - offset] = ['']
          }
      }
  
      ////varre os arrays multiplicando os valores
      for (i in arrayCustos.values){
          if (arred4(arrayCustos.values[i] * arrayCredICMS.values[i]) >= 0){
              creditoEmReais = creditoEmReais + arred4(arrayCustos.values[i] * arrayCredICMS.values[i]);
          }
      }
      despesas.getRange("D15").formulas = `${arred4(creditoEmReais)}`;

      // MARGEM LÍQUIDA
      despesas.getRange("D16").formulas = `=D3 + SUM(D4:D14) + D15`;

  });

}