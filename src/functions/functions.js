/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//const { Parser } = require("webpack");

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function getValueForKeyCF(key) {
  return getValueForKey(key);
}

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function setValueForKeyCF(key) {
  var value = 2;
  setValueForKey(key, value);
  console.log(key)
  return "Stored key/value pair";
}

/**
 * Get value for key
 * @customfunction
 * @param arrayCustos Custos
 * @returns The value for the key.
 */
 async function markup(num_linha, aliq) {
    //valida se algum campo não está selecionado na planilha
    var linha = await buscaLinha(num_linha);

    tipoItem = linha[1].toUpperCase();
    moeda = linha[14].toUpperCase();
    custoUnitario = linha[15];
    desconto = linha[18];
    txImp = linha[21];
    subTrib = linha[23];
    eImportado = linha[24];
    anexoIX = linha[25];
    credICMS = linha[26];


    if (tipoItem == '' || moeda == '' || subTrib == '' || eImportado == '' || anexoIX == '') return 0;

//    tipoItem = tipoItem.toUpperCase();
//    moeda = moeda.toUpperCase();
    subTrib = (subTrib.toUpperCase() == 'SIM');
    eImportado = (eImportado.toUpperCase() == 'SIM');
    anexoIX = (anexoIX.toUpperCase() == 'SIM');
    credICMS = credICMS * (!p.state.zerarICMS) * (tipoItem == 'HW' || tipoItem == 'MAT');

    var taxas = p.state.comDir + p.state.comCom + p.state.comPre + p.state.comPar
                    +  p.state.margem + p.state.txAdm + p.state.svTerc;
    
    var impostos = 0;
    var icms = calcICMS(tipoItem, eImportado, subTrib, anexoIX);
    var difal = p.state.ufOrig == p.state.ufDest ? 0: (icmsDaTabela(p.state.ufDest, p.state.ufDest) - icms) * !subTrib;

    var impostos = (() => {
        if (tipoItem == 'HW' || tipoItem == 'MAT'){
            return trib.state.csllHW + trib.state.irpjHW + trib.state.pis + trib.state.cofins + icms + difal;
        }else if (tipoItem == 'SW' || tipoItem == 'SV'){
            return trib.state.csllSW + trib.state.irpjSW + trib.state.pis + trib.state.cofins + trib.state.issOut * (!p.state.destGoiania) + trib.state.issGYN * (p.state.destGoiania);
        }else{
            //valida se o tipoItem está dentre os valores permitidos
            return -1;
        }
    });

    custoUnitario = (moeda == 'BRL' ? custoUnitario : ( moeda == 'USD' ? custoUnitario * p.state.dolarPTAX : (moeda == 'EUR'? custoUnitario * p.state.euroPTAX : 0)));
    custoComCredICMS = custoUnitario * (1 - credICMS)

    var taxaMarkup = 1/(1-(impostos() + taxas));

    
    //---------------------------------------
    console.log(`PARAMETROS\n
                DESCONTO: ${desconto}\n
                TX.IMP: ${txImp ==  0 ? 1 : txImp}\n
                CRED. ICMS: ${credICMS}\n
                TIPO ITEM: ${tipoItem}\n
                MOEDA: ${moeda}\n
                Subst. trib?: ${subTrib}\n
                IMPORTADO?: ${eImportado}\n
                ANEXO?: ${anexoIX}\n
                CUSTO EM REAIS: ${custoUnitario}\n
                CUSTO EM REAIS com cred: ${custoComCredICMS}\n
                `);
                
    console.log(`\nIMPOSTOS\n
                TOTAL: ${impostos()}\n
                CSLL HW: ${trib.state.csllHW}\n
                CSLL SW: ${trib.state.csllSW}\n
                IRPJ HW: ${trib.state.irpjHW}\n
                IRPJ SW: ${trib.state.irpjSW}\n
                PIS : ${trib.state.pis}\n
                COFINS : ${trib.state.cofins}\n
                ISS OUTROS: ${trib.state.issOut * (!p.state.destGoiania)}\n
                ISS GOIANIA: ${trib.state.issGYN * p.state.destGoiania}\n
                ICMS : ${icms}\n
                DIFAL : ${difal}\n
                `);

    console.log(`\nCOMISSÕES E TAXAS\n
                TOTAL: ${taxas + impostos()}\n
                `);
    console.log(`\nMARKUP\n
                TOTAL: ${taxaMarkup}\n
                `);
    //-------------------------------------------
    


    return arred2(((1-desconto)*(txImp ==  0 ? 1 : txImp)) * taxaMarkup * custoComCredICMS);
}

function icmsDaTabela(origem = p.state.ufOrig, destino = p.state.ufDest){
    return trib.state.tabelaIcms[listUF.indexOf(origem)][listUF.indexOf(destino)];
}

function calcICMS(tipoItem, eImportado = false, subTrib = false, anexoIX = false){
    if (p.state.zerarICMS) return 0;
    if (subTrib) return 0;
    if (tipoItem == 'SV' || tipoItem == 'SW' || tipoItem == '') return 0;

    if (eImportado && (p.state.ufOrig != p.state.ufDest)) {
        return 4/100;
    }else{
        if (p.state.ufOrig != p.state.ufDest){
            return icmsDaTabela()
        }else {
            if (anexoIX){
                return 7/100;
            }else{
                if (p.state.tipoFatur.toUpperCase() == 'GOVERNO'){
                    return 10.5/100;
                }else {
                    return icmsDaTabela();
                }
            }
        }
    }
}

function calcDIFAL(){
    var icms_destino = trib.state.tabelaIcms[listUF.indexOf(p.state.ufDest)][listUF.indexOf(p.state.ufDest)];
    var icms_inter = trib.state.tabelaIcms[listUF.indexOf(p.state.ufOrig)][listUF.indexOf(p.state.ufDest)];
    
    //essa multiplicação e divisão por 1000 é para arredondar para ter no máximo 3 casas decimais
    //sem essas operações, a função estava reornando um número 0.0800000000002
    return Math.trunc((icms_destino - icms_inter)*1000)/1000;

}

function porextenso(valor){
    //console.log(valor);
    
    var reais = valor.toString().split(".")[0];
    reais = reais != "0" ? reais : "";
    var centavos = valor.toString().split(".")[1];
    //essa linha transforma .1 em .10
    centavos = centavos ? centavos : "";
    centavos.length == 1 ? centavos = centavos.centavos = parseFloat(centavos).toPrecision(2).replace(".","") : 0;

    if (parseInt(reais) == 1){
        reais = reais.extenso();
        reais = reais.concat(" real");
    }else if (parseInt(reais) > 1) {
        reais = reais.extenso();
        reais = reais.concat(" reais");
    }

    if (parseInt(centavos) == 1){
        centavos = centavos.extenso();
        centavos = centavos.concat(" centavo");
    }else if (parseInt(centavos) > 1){
        centavos = centavos.extenso();
        centavos = centavos.concat(" centavos");
    }

    if (!centavos && !reais) {
        return ""
    }else if (!centavos){
        return reais
    }else if (!reais){
        return centavos
    }else{
        return reais + " e " + centavos
    }

}


function ptax(moeda){

    moeda = moeda.toUpperCase();

    if (moeda == 'USD'){
        return p.state.dolarPTAX;
    }else if (moeda == 'EUR'){
        return p.state.euroPTAX;
    }else if (moeda == 'BRL'){
        return 1
    }
    return -1;
}

async function buscaLinha(linha) {
    return Excel.run(async (context)=>{
      const ws = context.workbook.worksheets.getItem("Precificação");
      var  range = ws.getRange("B"+ linha + ":AB" + linha).load("values");
      await context.sync();
      return range.values[0];
    })
}

function impostos(tipo){
    var icms = icmsDaTabela(p.state.ufDest, p.state.ufDest) * (!p.state.zerarICMS);

    tipo = tipo.toUpperCase();

    if (tipo == 'HW' || tipo == 'MAT'){
        return trib.state.csllHW + trib.state.irpjHW + trib.state.pis + trib.state.cofins + icms;
    }else if (tipo == 'SW' || tipo == 'SV'){
        return trib.state.csllSW + trib.state.irpjSW + trib.state.pis + trib.state.cofins + trib.state.issOut * (!p.state.destGoiania) + trib.state.issGYN * (p.state.destGoiania);
    }else{
        //valida se o tipo está dentre os valores permitidos
        return -1;
    }
}



CustomFunctions.associate("GETVALUEFORKEYCF", getValueForKeyCF);
CustomFunctions.associate("SETVALUEFORKEYCF",setValueForKeyCF);
CustomFunctions.associate("MARKUP", markup);
CustomFunctions.associate("POREXTENSO", porextenso);
CustomFunctions.associate("PTAX", ptax);
CustomFunctions.associate("IMPOSTOS", impostos);