/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = () => {
  let keys = [];
  let values = [];

  // state object is used to track key/value pairs, and which storage type is in use
  g.state = {
    keys: keys,
    values: values,
    storageType: "globalvar"
  };

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
    irpjHW: 0,
    irpjSW: 0,
    csllHW: 0,
    csllSW: 0,
    cppHW: 0,
    cppSW: 0,
    issGYN: 0,
    issOut: 0,
    pis: 0,
    cofins: 0,
    tabelaIcms: []
  }
  

    // Connect handlers
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btnStoreValue").onclick = btnStoreValue;
    document.getElementById("btnGetValue").onclick = btnGetValue;
    document.getElementById("globalvar").onclick = btnStorageChanged;
    document.getElementById("localstorage").onclick = btnStorageChanged;
    document.getElementById("btnKit").onclick = novoKit;
    document.getElementById("btnTabela").onclick = novaTabela;
    document.getElementById("btnLinha").onclick = novaLinha;
    document.getElementById("btnRemoverLinha").onclick = removeLinha;
    document.getElementById("radioCambioData").onchange = radioListener;
    document.getElementById("radioCambioManual").onchange = radioListener;
    //document.getElementById("btnUpdateCambio").onclick = cambioListener;
    document.getElementById("inputDate").onchange = inputDateListener;
    document.getElementById("inputDate").onclick = inputDateListener;
    document.getElementById("inputUSD").onchange = inputCambioManualListener;
    document.getElementById("inputEUR").onchange = inputCambioManualListener;
    
    document.getElementById("inputMargem").onchange = atualizaMargem;
    document.getElementById("inputTxAdm").onchange = atualizaMargem;
    document.getElementById("inputSvTerc").onchange = atualizaMargem;
    document.getElementById("checkDiretoria").onchange = atualizaMargem;
    document.getElementById("checkComercial").onchange = atualizaMargem;
    document.getElementById("checkParceiro").onchange = atualizaMargem;
    document.getElementById("checkPrevendas").onchange = atualizaMargem;
    document.getElementById("inputTipoFaturamento").onchange = atualizaMargem;
    document.getElementById("inputUFOrigem").onchange = atualizaMargem;
    document.getElementById("inputUFDestino").onchange = atualizaMargem;
    document.getElementById("checkGoiania").onchange = atualizaMargem;
    document.getElementById("checkICMS").onchange = atualizaMargem;
    document.getElementById("btn-login").onclick = mostraOsPermitidos;
    document.getElementById("btn-add-sheet-sv").onclick = copiarPlanilhaSV;
    document.getElementById("btn-add-sheet-br").onclick = novaPlanilhaCustomizada;
    document.getElementById("btn-rem-sheet-sv").onclick = removePlanilhaSV;
    document.getElementById("btnContrib").onclick = calculaContribuicao;
    document.getElementById("btnCronograma").onclick = cronograma;
    document.getElementById("btnDI").onclick = copiaTabelaParaDI;
    document.getElementById("input-list-planilha").onkeyup = planilhaSV;

    document.getElementById("btn-mvr-sheet").onclick = moveParaDireita;
    document.getElementById("btn-mvl-sheet").onclick = moveParaEsquerda;
    //document.getElementById("a-implantacao").onclick = mostraImplantacao;

    //document.getElementById("input-list-planilha-br").onkeyup = planilhaBR;

    

    comeceAqui();
    //exibePlanilhas()

    var app_body = document.getElementById("app-body");
  
  linha = null;

};

/***
 * Handles the Store button press event and calls helper method to store the key/value pair from the user in storage.
 */
function btnStoreValue() {
  const keyElement = document.getElementById("txtKey");
  const valueElement = document.getElementById("txtValue");
  setValueForKey(keyElement.value, valueElement.value);
}

/***
 * Handles the Get button press and calls helper method to retrieve the value from storage for the given key.
 */
function btnGetValue() {
  const keyElement = document.getElementById("txtKey");
  (document.getElementById("txtValue")).value = getValueForKey(keyElement.value);
}

/***
 * Handles when the radio buttons are selected for local storage or global variable storage.
 * Updates a global variable that tracks which storage type is in use.
 */
function btnStorageChanged() {
  if ((document.getElementById("globalvar")).checked) {
    g.state.storageType = "globalvar";
  } else {
    g.state.storageType = "localstorage";
  }
}

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
      workbook.load("protection/protected");
  
      return context.sync().then(function() {
          if (!workbook.protection.protected) {
              workbook.protection.protect(SECRET);
          }
      });
      
    });
    

    
    //document.getElementById("perm-all-other-fields").hidden = true;
    //atualiza array com informações das tabelas
    //await atualizaArrayTabelas();
}


//atualiza da planilha para o frontend
function buscaMargem(){
  //document.getElementById('inputMargem').value = p.state.margem * 100;
  //document.getElementById('inputTxAdm').value = p.state.txAdm * 100;
  //document.getElementById('inputSvTerc').value = p.state.svTerc * 100;
  document.getElementById('inputMargem').value = (Math.round(p.state.margem*1000000)/10000);
  document.getElementById('inputTxAdm').value = (Math.round(p.state.txAdm*1000000)/10000);
  document.getElementById('inputSvTerc').value = (Math.round(p.state.svTerc*1000000)/10000);
  document.getElementById('checkDiretoria').checked  = p.state.aplicarComDir;
  document.getElementById('checkComercial').checked  = p.state.aplicarComCom;
  document.getElementById('checkPrevendas').checked  = p.state.aplicarComPre;
  document.getElementById('checkParceiro').checked  = p.state.aplicarComPar;
  
  
  document.getElementById('inputTipoFaturamento').value = p.state.tipoFatur;
  document.getElementById('inputUFOrigem').value = p.state.ufOrig;
  document.getElementById('inputUFDestino').value = p.state.ufDest;
  document.getElementById('checkGoiania').checked  = p.state.destGoiania;
  document.getElementById('checkICMS').checked  = p.state.zerarICMS;
  
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

    //console.log(`MARGEM: ${Math.round(p.state.margem * 10000)/100}%`)
    
    
    //calcula comissao diretoria
    if (!p.state.aplicarComDir) {
        //console.log("FALSE")
        p.state.comDir = 0;
    }else{
        //console.log(`MARGEM: ${Math.round(p.state.margem * 10000)/100}%`)
        if(p.state.margem < 0.1) {
            p.state.comDir = 0;
        }else if (p.state.margem >= 0.1 && p.state.margem <= 0.15){
            p.state.comDir = 0.005;
        }else {
            p.state.comDir = 0.01;
        }

    }
   // console.log(`COMIS. DIRETORIA: ${p.state.comDir * 100}%`)


    //calcula comissao comercial
    if (!p.state.aplicarComCom) {
        //console.log("FALSE")
        p.state.comCom = 0;
    }else{
        //console.log(`MARGEM: ${Math.round(p.state.margem * 10000)/100}%`)
        if(p.state.margem < 0.1) {
            p.state.comCom = 0.0015;
        }else if (p.state.margem >= 0.1 && p.state.margem <= 0.20){
            p.state.comCom = 0.01;
        }else if (p.state.margem > 0.20 && p.state.margem <= 0.30){
            p.state.comCom = 0.02;
        }else {
            p.state.comCom = 0.03;
        }
    }
    //console.log(`COMIS. COMERCIAL: ${p.state.comCom * 100}%`)

        //calcula comissao prevendas
        if (!p.state.aplicarComPre) {
          //console.log("FALSE")
          p.state.comPre = 0;
      }else{
          //console.log(`MARGEM: ${Math.round(p.state.margem * 10000)/100}%`)
          if(p.state.margem < 0.1) {
              p.state.comPre = 0.0015;
          }else if (p.state.margem >= 0.1 && p.state.margem <= 0.20){
              p.state.comPre = 0.0015;
          }else if (p.state.margem > 0.20 && p.state.margem <= 0.30){
              p.state.comPre = 0.0015;
          }else {
              p.state.comPre = 0.0015;
          }
      }  
      //console.log(`COMIS. PRÉ-VENDAS: ${p.state.comPre * 100}%`)
    
    //calcula comissao parceiro

    if (!p.state.aplicarComPar) {
        //console.log("FALSE")
        p.state.comPar = 0;
    }else{
        if(p.state.margem < 0.1) {
            p.state.comPar = 0;
        }else if (p.state.margem >= 0.1 && p.state.margem < 0.15){
            p.state.comPar = 0.05;
        }else if (p.state.margem >= 0.15 && p.state.margem < 0.20){
            p.state.comPar = 0.1;
        }else if (p.state.margem >= 0.20 && p.state.margem < 0.25){
            p.state.comPar = 0.15;
        }else {
            p.state.comPar = 0.2;
        }
    }  
    //console.log(`COMIS. PARCEIRO: ${p.state.comPar * 100}%`)

    if (p.state.ufDest != 'GO') {
      document.getElementById('checkGoiania').disabled = true;
      document.getElementById('checkGoiania').checked  = false;
      p.state.destGoiania = false;
      //checkGoiania
    } else {
      document.getElementById('checkGoiania').disabled = false;
    }
    
    await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem("param");
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
    });
    
}


async function atualizaTributos() {
  Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem("trib");
    var  range = ws.getRange("A1:B22").load("values");
    await context.sync();

    trib.state.irpjHW = range.m_values[1][1];
    trib.state.irpjSW = range.m_values[2][1];
    trib.state.csllHW = range.m_values[5][1];
    trib.state.csllSW = range.m_values[6][1];
    trib.state.cppHW = range.m_values[9][1];
    trib.state.cppSW = range.m_values[10][1];
    trib.state.issGYN = range.m_values[13][1];
    trib.state.issOut = range.m_values[14][1];
    trib.state.pis = range.m_values[17][1];
    trib.state.cofins = range.m_values[21][1];

    trib.state.tabelaIcms = ws.getRange("F3:AF29").load("values");
    await context.sync();

    trib.state.tabelaIcms = trib.state.tabelaIcms.m_values;


    //                            destino        x      origem
    //trib.state.tabelaIcms[listUF.indexOf("PB")][listUF.indexOf("RJ")]

    //console.log( trib.state.tabelaIcms);

    return context;
  })
}

async function atualizaParametros() {
  Excel.run(async (context)=>{
    const ws = context.workbook.worksheets.getItem("param");
    var  range = ws.getRange("B1:B21").load("values");
    await context.sync();

    p.state.hoje = new Date();    
    
    p.state.tipoFatur     = range.m_values[0][0];
    p.state.margem        = range.m_values[1][0];
    p.state.comDir        = range.m_values[2][0];
    p.state.comCom        = range.m_values[3][0];
    p.state.comPre        = range.m_values[4][0];
    p.state.comPar        = range.m_values[5][0];
    p.state.txAdm         = range.m_values[6][0];
    p.state.svTerc        = range.m_values[7][0];
    p.state.aplicarComDir = range.m_values[8][0];
    p.state.aplicarComCom = range.m_values[9][0];
    p.state.aplicarComPre = range.m_values[10][0];
    p.state.aplicarComPar = range.m_values[11][0];
    p.state.ufOrig        = range.m_values[12][0];
    p.state.ufDest        = range.m_values[13][0];
    p.state.destGoiania   = range.m_values[14][0];
    p.state.zerarICMS     = range.m_values[15][0];
    p.state.txImpHW       = range.m_values[16][0];
    p.state.txImpSW       = range.m_values[17][0];
    p.state.dataCambio    = ExcelDateToJSDate(range.m_values[18], -3);
    p.state.dolarPTAX     = range.m_values[19];
    p.state.euroPTAX      = range.m_values[20];

    

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

async function ocultaPlanilhas() {
  Excel.run(async (context)=>{
    const ws = context.workbook.worksheets;
    ws.getItem("trib").visibility = Excel.SheetVisibility.veryHidden ;
    ws.getItem("param").visibility = Excel.SheetVisibility.veryHidden ;
    ws.getItem("list").visibility = Excel.SheetVisibility.veryHidden ;
    ws.getItem("modelos").visibility = Excel.SheetVisibility.veryHidden ;
    ws.getItem("login").visibility = Excel.SheetVisibility.veryHidden ;
    
    await context.sync();

    return context;
  })
}

async function exibePlanilhas() {
  Excel.run(async (context)=>{
    const ws = context.workbook.worksheets;
    ws.getItem("param").visibility = Excel.SheetVisibility.visible ;
    ws.getItem("modelos").visibility = Excel.SheetVisibility.visible
    ws.getItem("trib").visibility = Excel.SheetVisibility.visible ;
    ws.getItem("list").visibility = Excel.SheetVisibility.visible ;
    ws.getItem("login").visibility = Excel.SheetVisibility.visible ;
    
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
  console.log(perfil)
  return permissoes;
}


//retorna o código de permissão (10, 20, 30, 40 ou 50). Se der erro, retorna -1.
async function getProfile(login = document.getElementById('input-passwd').value) {
  var base = await getAllLogin();
  var users = [];
  var permis = [];

  for (i in base){
    users[i] = base[i][0];
    permis[i] = base[i][1];
  }
  
  let perm = permis[users.indexOf(login)];

  if (typeof perm == 'undefined' || perm == ''){
    return -1;
  }
  console.log(login)
  return perm;
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

async function mostraOsPermitidos(){
  //document.getElementById("perm-login").hidden = true;
  perfil = await getProfile();
  var permissoes = await getPermissions(perfil);

  if(perfil == -1){
    console.log("usuário inválido")
    return -1;
  }
  
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
      case 'cambio':
        link.text = 'Implantação'
        break;
        
      default:
        break;
    }

    listItem.id = 'li-'+nome;
    listItem.className = 'nav-item';
    listItem.appendChild(link);

    return listItem;

  })




  document.getElementById('nav-ul').appendChild(li('edicao'));
  document.getElementById('nav-ul').appendChild(li('cambio'));
  document.getElementById('nav-ul').appendChild(li('margens'));
  document.getElementById("a-cambio").onclick = mostraCambio;
  document.getElementById("a-edicao").onclick = mostraEdicao;
  document.getElementById("a-margens").onclick = mostraMargens;


  document.getElementById("perm-all-other-fields").hidden = false;
  document.getElementById("nav-ul").hidden = false;
  document.getElementById("perm-login").hidden = true;

  // document.getElementById("perm-cambio").hidden = !permissoes[1];
  // document.getElementById("perm-faturamento").hidden = !permissoes[2];
  // document.getElementById("perm-margens").hidden = !permissoes[3];
  // document.getElementById("perm-comissoes").hidden = !permissoes[3];
  // document.getElementById("perm-formatar").hidden = !permissoes[10];

  // document.getElementById("radioCambioManual").disabled = permissoes[14];
  // document.getElementById("radioCambioData").disabled = permissoes[14];
  // document.getElementById("inputDate").disabled = permissoes[14];
  // document.getElementById("inputUSD").disabled = permissoes[14];
  // document.getElementById("inputEUR").disabled = permissoes[14];

  return permissoes;

// legenda de permissões
//
//00 ADM
//01 CAMBIO
//02 FATURAMENTO
//03 MARGEM
//04 TXADM
//05 SVTERC
//06 DIRETORIA
//07 COMERCIAL
//08 PARCEIRO
//09 PREVENDAS
//10 FORMATAR PLANILHAS
//11 RESUMO
//12 SIECON
//13 CRONOGRAMA
//14 CAMBIO-LEITURA

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
  const input = document.getElementById("input-list-planilha").value;
  const datalist = document.getElementById("datalist-planilha").options;
  //console.log(inputSV.value);

  if (input == "SV" || input == "param" || input == "list" || input == "login" || input == "trib" ||  input == "extenso" ||  input == "Precificação" || input == "CRONOGRAMA-COMPRAS" || input == "ANEXO IV" || input == "modelos" || input == "DESPESAS-INDIRETAS" ){
    document.getElementById("btn-add-sheet-sv").disabled = true;
    document.getElementById("btn-add-sheet-br").disabled = true;
    document.getElementById("btn-rem-sheet-sv").disabled = true;
    return
  }

  if (input == ''){
    document.getElementById("btn-add-sheet-sv").disabled = false;
    document.getElementById("btn-add-sheet-br").disabled = false;
    document.getElementById("btn-rem-sheet-sv").disabled = true;
  }


  if (existeNaLista(input)){
    document.getElementById("btn-add-sheet-sv").disabled = true;
    document.getElementById("btn-add-sheet-br").disabled = true;
    document.getElementById("btn-rem-sheet-sv").disabled = false;
  }else{
    document.getElementById("btn-add-sheet-sv").disabled = false;
    document.getElementById("btn-add-sheet-br").disabled = false;
    document.getElementById("btn-rem-sheet-sv").disabled = true;
  }

}

//procura se um valor existe na lista de planilhas SV ou BR
//retorna true se existe e false se não
function existeNaLista (valor){
  const datalist = document.getElementById("datalist-planilha").options;

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
    document.getElementById('datalist-planilha').appendChild(node)
  }else{
    console.log("já existe")
  }
}


function mostraCambio(){
  //oculta todas os campos
  //oculta login
  document.getElementById("perm-login").hidden = true;
  //oculta faturamento
  document.getElementById("perm-faturamento").hidden = false;
  //oculta margens
  document.getElementById("perm-margens-comissoes").hidden = true;

  //oculta formatar
  document.getElementById("perm-formatar").hidden = true;
  //oculta implantacao
  document.getElementById("perm-implantacao").hidden = true;
  //exibe o campo de PTAX
  document.getElementById("perm-cambio").hidden = false;

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
//       console.log(result.value)
//     }
//   );
// }

// function processMessage(arg) {
//   document.getElementById("user-name").innerHTML = arg.message;
//   console.log(arg.message)
//   dialog.close();
// }
// var dialog = null;