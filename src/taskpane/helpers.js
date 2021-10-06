/* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. */


/***
 * Stores the key/value pair. Will use local storage or global variable to store
 * the values depending on which type the user selected.
 * 
 * @export
 * @param {string} key The key to store.
 * @param {string} value The value to store.
 */
function setValueForKey(key, value) {
    if (g.state.storageType === "globalvar") {
      g.state.keys.push(key);
      g.state.values.push(value);
    } else {
      window.localStorage.setItem(key, value);
    }
}

/**
 * Gets the value for the given key from storage. Will retrieve the value
 * from local storage or global variable depending on which type of storage
 * the user selected.
 *
 * @export
 * @param {string} key The key to retrieve the value for
 * @returns {string} The value
 */
function getValueForKey(key) {
    let answer = "";
    if (g.state.storageType === "globalvar") {
      // get value from global variable
      g.state.keys.forEach((element, index) => {
        if (element === key) {
          answer = g.state.values[index];
        }
      });
    } else {
      // get value from localStorage
      answer = window.localStorage.getItem(key);
    }
    return answer;
}


//essa função é ativada sempre que o radio for alterado (onchange)
function radioListener(){
    //se selecionar o radio da data, deve desabilitar os campos de input manual 
    if (document.getElementById("radioCambioData").checked == true){
        document.getElementById("inputDate").disabled = false;
        //document.getElementById("btnUpdateCambio").disabled = false;
        document.getElementById("inputUSD").disabled = true;
        document.getElementById("inputEUR").disabled = true;
        document.getElementById("inputDate").value = ontemStringHTML();
        cambioListener();
    }
    //se selecionar o radio manual, deve desabilitar o campo e botão de atualizar
    if (document.getElementById("radioCambioManual").checked == true){
        document.getElementById("inputDate").disabled = true;
        //document.getElementById("btnUpdateCambio").disabled = true;
        document.getElementById("inputUSD").disabled = false;
        document.getElementById("inputEUR").disabled = false;
        document.getElementById("inputDate").value = "";
    }
}


function ExcelDateToStrDate(serialDate, offsetUTC) {
  // serialDate is whole number of days since Dec 30, 1899
  // offsetUTC is -(24 - your timezone offset)
  // eu acrescentei o "-18" abaixo para que a função seja chamada com "-3"
  var jsdate = new Date(Date.UTC(0, 0, serialDate, -18 + offsetUTC));
  return jsdate.toISOString().split('T')[0];
}

function JSDateToStrDate(jsdate) {
  return jsdate.toISOString().split('T')[0];
}

function ExcelDateToJSDate(serialDate, offsetUTC) {
  // serialDate is whole number of days since Dec 30, 1899
  // offsetUTC is -(24 - your timezone offset)
  // eu acrescentei o "-18" abaixo para que a função seja chamada com "-3"
  return new Date(Date.UTC(0, 0, serialDate, -18 + offsetUTC));
}

function ontemStringHTML(){
  var ontem = new Date();
  ontem.setDate(ontem.getDate() - 1);
  return JSDateToStrDate(ontem)
}


//essa função busca a cotação da moeda em um determinado dia e em um intervalo de dias atrás.
//ela retorna a cotação mais próxima da data passada como parâmetro.
//o intervalo é necessário pois em dias não comerciais, não há cotação. Para evitar que
//a busca retorne vazia, o intervalo garante que haverá cotações válidas.
//function buscaPTAX(data, moeda, intervalo = 5){
async function buscaPTAX(moeda, dataFinal = ontemStringHTML(), intervalo = 5){
    
    let obj = null;
    var dataInicial = calculaDataInicial (dataFinal, intervalo);
    dataInicial = dataHTMLparaEUA(dataInicial);
    dataFinal = dataHTMLparaEUA(dataFinal);
    let url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?@moeda='" + moeda + "'&@dataInicial='" + dataInicial + "'&@dataFinalCotacao='" + dataFinal + "'&$top=100&$filter=tipoBoletim%20eq%20'Fechamento'&$orderby=dataHoraCotacao%20desc&$format=json&$select=cotacaoVenda,dataHoraCotacao,tipoBoletim";

    try {
        obj = await (await fetch(url)).json();
    } catch(e) {
        console.log('error');
        console.log(url);
    }
        
    return obj.value[0].cotacaoVenda;
    //return obj;
    
}

//o objeto input date do frontend retorna value no formato YYYY-MM-DD
//essa função converte para formato EUA: MM-DD-YYY, que é aceito pela API do BC
function dataHTMLparaEUA(dataHTML){
    dataHTML = dataHTML.split('-');
    dataHTML.push(dataHTML[0]);
    dataHTML.shift();
    return dataHTML.join('-');
}

//retorna uma string no formato YYY-MM-DD
//o valor da data retornado é (dataFinal - intervalo)
function calculaDataInicial (dataFinal, intervalo){
    let df = dataFinal.split('-').map(Number);
    let date = new Date(df[0], df[1] - 1, df[2]);
    date.setDate(date.getDate() - intervalo);
    return date.toISOString().split('T')[0];
}


//busca valor do cambio no BCB e atualiza a planilha e o frontend
async function cambioListener (){

    //atualizaParametros();
    
    //se não há data selecionada no campo de data, usa a data de ontem
    if (document.getElementById("inputDate").value == "") {
        document.getElementById("radioCambioData").checked = true;
        radioListener();
        
        //document.getElementById("inputDate").value = ontemStringHTML();
    }

    //busca as cotações no BCB
    p.state.dataCambio = document.getElementById("inputDate").value
    p.state.dolarPTAX = await buscaPTAX("USD", document.getElementById("inputDate").value);
    p.state.euroPTAX = await buscaPTAX("EUR", document.getElementById("inputDate").value);
    document.getElementById("inputUSD").value = p.state.dolarPTAX;
    document.getElementById("inputEUR").value = p.state.euroPTAX;

    //salva os valores e data na planilha
    Excel.run (async(context) =>{
        var ws = context.workbook.worksheets.getItem("param");
        var range = ws.getRange("B19:B21").load("values");
        //var range = ws.getRange("B16:B18");
        await context.sync();
        let novosvalores = [[document.getElementById("inputDate").value], [document.getElementById("inputUSD").value], [document.getElementById("inputEUR").value]];
        
        
        //await context.sync();
        range.values = novosvalores;
        return context.sync();
    })

    //atualizaParametros();
    
}


//atualiza o frontend com as informações da planilha e parametros globais p.state
function atualizaDivPTAX(){

    // set o valor máximo = ontem
    document.getElementById("inputDate").max = ontemStringHTML();

    // set valor atual e seleciona o radio
    try{
        if (p.state.dataCambio.getFullYear() == 1899){
            document.getElementById("inputDate").value = "";
            document.getElementById("radioCambioManual").checked = true;
            document.getElementById("inputUSD").value = p.state.dolarPTAX;
            document.getElementById("inputEUR").value = p.state.euroPTAX;
        } else {
            document.getElementById("inputDate").value = JSDateToStrDate(p.state.dataCambio);
            document.getElementById("radioCambioData").checked = true;
            document.getElementById("inputUSD").value = p.state.dolarPTAX;
            document.getElementById("inputEUR").value = p.state.euroPTAX;
        }
    }catch(e){
        console.log("Parâmetro não carregado.");
    }

    //se selecionar o radio da data, deve desabilitar os campos de input manual 
    if (document.getElementById("radioCambioData").checked == true){
        document.getElementById("inputDate").disabled = false;
        //document.getElementById("btnUpdateCambio").disabled = false;
        document.getElementById("inputUSD").disabled = true;
        document.getElementById("inputEUR").disabled = true;

    }
    //se selecionar o radio manual, deve desabilitar o campo e botão de atualizar
    if (document.getElementById("radioCambioManual").checked == true){
        document.getElementById("inputDate").disabled = true;
        //document.getElementById("btnUpdateCambio").disabled = true;
        document.getElementById("inputUSD").disabled = false;
        document.getElementById("inputEUR").disabled = false;
    }
    
}

async function inputDateListener() {
    
    //document.getElementById("inputUSD").value = await buscaPTAX("USD", document.getElementById("inputDate").value);
    //document.getElementById("inputEUR").value = await buscaPTAX("EUR", document.getElementById("inputDate").value);
    cambioListener ();

}

async function inputCambioManualListener() {

    p.state.dataCambio = ExcelDateToJSDate("", -3);
    //p.state.dolarPTAX = await buscaPTAX("USD", document.getElementById("inputDate").value);
    //p.state.euroPTAX = await buscaPTAX("EUR", document.getElementById("inputDate").value);
    p.state.dolarPTAX = document.getElementById("inputUSD").value
    p.state.euroPTAX = document.getElementById("inputEUR").value;
    document.getElementById("inputDate").value = "";

    //salva os valores e data na planilha
    Excel.run (async(context) =>{
        var ws = context.workbook.worksheets.getItem("param");
        var range = ws.getRange("B19:B21").load("values");
        //var range = ws.getRange("B16:B18");
        await context.sync();
        let novosvalores = [[""], [document.getElementById("inputUSD").value], [document.getElementById("inputEUR").value]];
        
        //await context.sync();
        range.values = novosvalores;
        
        return context.sync();
    })
    

}

async function novaLinha(){
    const numLinhas = parseInt(document.getElementById("inputNumLinha").value);
    await atualizaArrayTabelas();
    await Excel.run(async (context) =>{

        const cell = context.workbook.getActiveCell();
        const a = cell.getCellProperties({address: true});
        await context.sync()
        
        if (a.value[0][0]["address"].split('!')[0] != 'Precificação'){
            console.log('Not in sheet');
            return 0
        }

        const ws = context.workbook.worksheets.getItem("Precificação");
        var index_tabela = await estaEmTabela();

        if (index_tabela == -1){
            console.log('Not in sheet');
            return "Fora de tabela";
        }
        
        ws.getRange(tabelas[index_tabela -1].linha_fin.toString().concat(":"+ (tabelas[index_tabela -1].linha_fin + numLinhas -1 ))).insert(Excel.InsertShiftDirection.down);
        ws.getRange(tabelas[index_tabela -1].linha_fin.toString().concat(":"+ (tabelas[index_tabela -1].linha_fin + numLinhas -1 ))).copyFrom("modelos!4:4");
        tabelas[index_tabela - 1].linha_fin += numLinhas ;

        //ATUALIZA A FORMULA SUBTOTAL (OBS: AQUI DEVE SER ESCRITA A FORMULA COMO EXCEL EM INGLES!! NÃO USAR PONTO-E-VIRGULA E NOMES EM PT-BR)
        ws.getRange("K"+ tabelas[index_tabela - 1].linha_fin).formulas =
                            [["=SUBTOTAL(9,K" +( tabelas[index_tabela - 1].linha_ini + 2) + ":K" + (tabelas[index_tabela - 1].linha_fin - 1) + ")"]];
         await context.sync();
        
    })
    await atualizaArrayTabelas();
    await renumerar();

}

async function setFormula(formula, address) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Precificação");
  
      const range = sheet.getRange(address);
      range.formulas = [[formula]];
      //range.format.autofitColumns();
  
      await context.sync();
    });
  }

async function getLastCellAddress() {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Precificação");
        const range = sheet.getUsedRange();
        var obj = {};
        range.load("address");
  
        await context.sync();
        obj.linha = parseInt(range.address.split(':')[1].replace(/[A-Z]/g, ''));
        obj.coluna = range.address.split(':')[1].replace(/[0-9]/g, '');
        //console.log(`The address of the used range in the worksheet is "${range.address}"`);
        return obj;
    });
  }

async function copyAll() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Precificação");
      // Place a label in front of the copied data.
      //sheet.getRange("F1").values = [["Copied Range"]];
  
      // Copy a range starting at a single cell destination.
      //.insert(Excel.InsertShiftDirection.down)
      sheet.getRange("36:38").insert(Excel.InsertShiftDirection.down);//copyFrom("modelos!1:9");
      await context.sync();
    });
  }


//retorna a linha da célula selecionada na planilha.
//caso a seleção seja um range de células, retorna a linha  da célula superior esquerda[0] e a planilha[1]
async function linhaSelecionada(){

    //, activeCell.address.split('!')[0].replace(/[0-9]/g, '')
    const endereco = await Excel.run(async (context) =>{

        const activeCell = context.workbook.getActiveCell();
        activeCell.load("address");

        await context.sync();
        return [activeCell.address.split('!')[1].replace(/[A-Z]/g, '') , activeCell.address.split('!')[0].replace(/[0-9]/g, '')];
    })

    endereco[0] = parseInt(endereco[0]);

    return endereco

}


//verifica se a linha selecionada está contida em alguma tabela de precificação.
//case positivo, retorna o index da tabela começando em 1.
//caso esteja fora da planilha ou fora de tabela, retorna -1
async function estaEmTabela(){

    let linha =  await linhaSelecionada();
    
    if (linha[1] == "Precificação"){
        for (i in tabelas) {
            if (tabelas[i].linha_ini <= linha[0]  && tabelas[i].linha_fin >= linha[0]){
                return tabelas[i].index;
            }
        }
    }
    return -1;
}

//atualiza array tabelas, que é variável global com as informações
// das tabelas de precificação
async function atualizaArrayTabelas() {
    //pega a ultima linha da aba precificação e salva em variavel
    const ultimaCelula = await getLastCellAddress();
    var celula = {};
    tabelas = [];

    ultimaCelula.coluna = colunas[ colunas["length"] -1 ].fin;

    //copia as propriedades da última coluna para a constante propriedades
    const propriedades = await getCellProperties(ultimaCelula.coluna.concat(1,':', ultimaCelula.coluna, ultimaCelula.linha));
    var numTabelas = 0;
    var indexTabela = 0;
    //console.log(ultimaCelula.coluna);
    

    var infoTabela = { "index":0 , "num_linhas":0 , "linha_ini":0 , "linha_fin":0 };

    //varre da primeira até a ultima linha da última coluna
    for(let i = 0; i < ultimaCelula.linha; i++){

        celula = propriedades[i][0];
        //se for uma célula com cor #2B2E34, significa que é uma nova tabela
        if (celula.format.fill.color == "#2B2E34"){

            numTabelas ++;
            tabelas.push({});
            tabelas[numTabelas -1].index = numTabelas;
            tabelas[numTabelas -1].linha_ini = i + 1;
            tabelas[numTabelas -1].num_linhas = 0;         

        } 
        //conta as células internas (com itens) da tabela
        if (! celula.style.includes("Normal")){
            //console.log(celula.address);
            //console.log(celula.style);
            tabelas[numTabelas -1 ].num_linhas ++;
            tabelas[numTabelas -1 ].linha_fin = i + 2;
        }
    }

    return ultimaCelula; 
}


            
//deve receber uma string no formato AA10, como as colunas no excel
//retorna objeto com as propriedades: .style .format.fill.color .address
async function getCellProperties(address) {
    
    return await Excel.run(async (context) => {
        
        
      const cell = context.workbook.worksheets.getItem("Precificação").getRange(address);
      // Define the cell properties to get by setting the matching LoadOptions to true.
      const propertiesToGet = cell.getCellProperties({
          address: true,
        format: {
          fill: {
            color: true
          },
          font: {
            color: true
          }
        },
        style: true
      });
  
      // Sync to get the data from the workbook.
      await context.sync()
      //return propertiesToGet.value[0][0].style;
      const cellProperties = propertiesToGet.value;
      return cellProperties;
      
    });
  }


  async function novaTabela() {
      await atualizaArrayTabelas()
      
      //console.log(tabelas.length);

      if (tabelas.length > 0){
          var linhaInicioNovaTabela =  tabelas[tabelas.length -1].linha_fin + 1;
      }else{
          var linhaInicioNovaTabela =  8;
      }
      //console.log(linhaInicioNovaTabela);
      let linStr = linhaInicioNovaTabela.toString().concat(":"+ (linhaInicioNovaTabela + 8));

      //return linStr;

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Precificação");

        sheet.getRange(linStr).insert(Excel.InsertShiftDirection.up);
        sheet.getRange(linStr).copyFrom("modelos!1:9");
        await context.sync();
      });
      await atualizaArrayTabelas();           
      await renumerar(); 
  }

  async function novoKit() {
      await atualizaArrayTabelas();
      var indexTabela = await estaEmTabela();

      if (indexTabela == -1) {
        console.log("fora de tabela");
        return -1
      }

      return await Excel.run(async (context) => {

        var range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();

        //return 0;
        
        //console.log(range.address)
        var kit = {
            "linha_ini": 0,
            "linha_fin": 0,
            "planilha": "",
        };

        kit.linha_ini = parseInt(range.address.split('!')[1].split(':')[0].replace(/[A-Z]/g, ''));
        kit.planilha = range.address.split('!')[0];

        //caso só tenha uma linha selecionada, daria erro no split
        //por isso coloquei o try-catch
        try{

            kit.linha_fin = parseInt(range.address.split('!')[1].split(':')[1].replace(/[A-Z]/g, ''));

        }catch (e){

            kit.linha_fin = parseInt(range.address.split('!')[1].replace(/[A-Z]/g, ''));

        }


        //valdar se o range está contido em alguma tabela:
        //1 - identificar se a primeira linha está em alguma tabela e qual o index dessa tabela
        //2 - identificar se a ultima linha do range esta na mesma tabela                    
        //a primeira linha do range do kit deve ser pelo menos a segunda linha da tabela (1.2) e no máximo a penúltima linha (pois a última é o subtotal)
        if ((tabelas[indexTabela - 1].linha_ini + 3) <= kit.linha_ini  
            && (tabelas[indexTabela - 1].linha_fin -1 ) >= kit.linha_ini
            && (tabelas[indexTabela - 1].linha_fin -1 ) >= kit.linha_fin ){
                //console.log("Seleção OK");
                //agrupa as células
                range.group(Excel.GroupOption.byRows);
                await context.sync();

                //seleciona demais células do kit e muda a formatação
                for (i in colunas){

                    var subitens = context.workbook.worksheets.getItem("Precificação").getRange(
                        colunas[i].ini + kit.linha_ini.toString() + ":" +
                        colunas[i].fin + kit.linha_fin.toString()  );
                    subitens.format.fill.color = "#D9D9D9";
                    subitens.format.borders.getItem('EdgeBottom').color = "#000000";
                    subitens.format.borders.getItem('EdgeRight').color = "#000000";
                    subitens.format.borders.getItem('EdgeLeft').color = "#000000";
                    subitens.format.borders.getItem('EdgeTop').color = "#F2F2F2";
                    subitens.format.borders.getItem('InsideHorizontal').color = "#F2F2F2";
                    subitens.format.borders.getItem('InsideVertical').color = "#000000";
                }              

                //FORMATAÇÃO DO HEADER
                for (i in colunas){
                    if (i > 0){
                        var head = context.workbook.worksheets.getItem("Precificação").getRange(
                            colunas[i].ini + (kit.linha_ini - 1).toString() + ":" +
                            colunas[i].fin + (kit.linha_ini - 1).toString()  );
                        
                        head.format.font.color = "#203764";
                        head.format.font.bold = true;
                    }   
                }

                //ARRUMA A borda superior da abaixo do kit---------------------------------
                for (i in colunas){
                    var subitens = context.workbook.worksheets.getItem("Precificação").getRange(
                        colunas[i].ini + (kit.linha_fin + 1).toString() + ":" +
                        colunas[i].fin + (kit.linha_fin + 1).toString()  );
                    subitens.format.borders.getItem('EdgeTop').color = "#000000";
                }

                //formulas do cabeçalho do kit
                //valor de venda unitario
                var celula = context.workbook.worksheets.getItem("Precificação").getRange(
                   "J" + (kit.linha_ini - 1).toString() + ":" +
                   "J" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror(SUBTOTAL(9,K${kit.linha_ini}:K${kit.linha_fin})/qtde,0)`]];

                //valor de venda total do item
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "K" + (kit.linha_ini - 1).toString() + ":" +
                    "K" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror(SUBTOTAL(9,K${kit.linha_ini}:K${kit.linha_fin})/qtde,0)*qtde`]];

                //moeda
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "P" + (kit.linha_ini - 1).toString() + ":" +
                    "P" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];
                
                //valor custo unitario
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "Q" + (kit.linha_ini - 1).toString() + ":" +
                    "Q" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror(SUBTOTAL(9,R${kit.linha_ini}:R${kit.linha_fin})/qtde,0)`]];
                
                //valor custo total
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "R" + (kit.linha_ini - 1).toString() + ":" +
                    "R" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror((SUBTOTAL(9,R${kit.linha_ini}:R${kit.linha_fin})/qtde)*qtde,0)`]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "T" + (kit.linha_ini - 1).toString() + ":" +
                    "T" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "V" + (kit.linha_ini - 1).toString() + ":" +
                    "V" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "W" + (kit.linha_ini - 1).toString() + ":" +
                    "W" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "Y" + (kit.linha_ini - 1).toString() + ":" +
                    "Y" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "Z" + (kit.linha_ini - 1).toString() + ":" +
                    "Z" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AA" + (kit.linha_ini - 1).toString() + ":" +
                    "AA" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AB" + (kit.linha_ini - 1).toString() + ":" +
                    "AB" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];
                
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AE" + (kit.linha_ini - 1).toString() + ":" +
                    "AE" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror(SUBTOTAL(9,AF${kit.linha_ini}:AF${kit.linha_fin})/qtde,0)`]];
                                
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AF" + (kit.linha_ini - 1).toString() + ":" +
                    "AF" + (kit.linha_ini - 1 ).toString());
                    celula.formulas = [[`=iferror((SUBTOTAL(9,AF${kit.linha_ini}:AF${kit.linha_fin})/qtde)*qtde,0)`]];
                                
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AH" + (kit.linha_ini - 1).toString() + ":" +
                    "AH" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror(SUBTOTAL(9,AI${kit.linha_ini}:AI${kit.linha_fin})/qtde,0)`]];

                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AI" + (kit.linha_ini - 1).toString() + ":" +
                    "AI" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[`=iferror((SUBTOTAL(9,AI${kit.linha_ini}:AI${kit.linha_fin})/qtde)*qtde,0)`]];
                
                celula = context.workbook.worksheets.getItem("Precificação").getRange(
                    "AK" + (kit.linha_ini - 1).toString() + ":" +
                    "AK" + (kit.linha_ini - 1 ).toString());
                celula.formulas = [[""]];
                await context.sync()

            await renumerar();
            return kit;
        }


        //sheet.getRange(linStr).insert(Excel.InsertShiftDirection.down);
        //sheet.getRange(linStr).copyFrom("modelos!1:9");
        console.log("fora de tabela")
        return kit;
      });

      console.log("dentro tabela");
  }

  function carregaListaUF() {
      const combo1 = document.getElementById("inputUFOrigem");
      const combo2 = document.getElementById("inputUFDestino");

      let option = document.createElement("option");

      for (i in listUF) {
        let option = document.createElement("option");
        option.value = listUF[i];
        option.text = listUF[i];

        //combo1.appendChild(option);
        combo2.appendChild(option);
      }
      

      for (i in listUF) {
        let option = document.createElement("option");
        option.value = listUF[i];
        option.text = listUF[i];

        combo1.appendChild(option);
        //combo2.appendChild(option);
      }
      combo1.value = "";
      combo2.value = "";
  }

//+ Carlos R. L. Rodrigues
//@ http://jsfromhell.com/string/extenso [rev. #3]
String.prototype.extenso = function(c){
    var ex = [
        ["zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"],
        ["dez", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"],
        ["cem", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"],
        ["mil", "milhão", "bilhão", "trilhão", "quadrilhão", "quintilhão", "sextilhão", "setilhão", "octilhão", "nonilhão", "decilhão", "undecilhão", "dodecilhão", "tredecilhão", "quatrodecilhão", "quindecilhão", "sedecilhão", "septendecilhão", "octencilhão", "nonencilhão"]
    ];
    var a, n, v, i, n = this.replace(c ? /[^,\d]/g : /\D/g, "").split(","), e = " e ", $ = "real", d = "centavo", sl;
    for(var f = n.length - 1, l, j = -1, r = [], s = [], t = ""; ++j <= f; s = []){
        j && (n[j] = (("." + n[j]) * 1).toFixed(2).slice(2));
        if(!(a = (v = n[j]).slice((l = v.length) % 3).match(/\d{3}/g), v = l % 3 ? [v.slice(0, l % 3)] : [], v = a ? v.concat(a) : v).length) continue;
        for(a = -1, l = v.length; ++a < l; t = ""){
            if(!(i = v[a] * 1)) continue;
            i % 100 < 20 && (t += ex[0][i % 100]) ||
            i % 100 + 1 && (t += ex[1][(i % 100 / 10 >> 0) - 1] + (i % 10 ? e + ex[0][i % 10] : ""));
            s.push((i < 100 ? t : !(i % 100) ? ex[2][i == 100 ? 0 : i / 100 >> 0] : (ex[2][i / 100 >> 0] + e + t)) +
            ((t = l - a - 2) > -1 ? " " + (i > 1 && t > 0 ? ex[3][t].replace("ão", "ões") : ex[3][t]) : ""));
        }
        a = ((sl = s.length) > 1 ? (a = s.pop(), s.join(" ") + e + a) : s.join("") || ((!j && (n[j + 1] * 1 > 0) || r.length) ? "" : ex[0][0]));
        a && r.push(a + (c ? (" " + (v.join("") * 1 > 1 ? j ? d + "s" : (/0{6,}$/.test(n[0]) ? "de " : "") + $.replace("l", "is") : j ? d : $)) : ""));
    }
    return r.join(e);
}


async function renumerar (){
    //atualizar index tabelas
    await atualizaArrayTabelas();

    //criar array vazio para numeros
    var numeracao = [[]];
    var celula = {};
    var linhaAtual = 0;
    var item = 1;
    var subitem = 1;
    
    //copia as propriedades da primeira coluna para a constante propriedades
    const coluna = "B";
    const primeiraLinha = tabelas[0].linha_ini;
    const ultimaLinha = tabelas[tabelas.length-1].linha_fin;

    const propriedades = await getCellProperties(coluna.concat(primeiraLinha, ':', coluna, ultimaLinha));
    
    if (tabelas.length == 0){
        return -1;
    }
    var indexTabela = 0
    
    //varrer coluna B verificando a formatação das células
    for(i in propriedades){
        celula = propriedades[i][0];
        celula.color = celula.format.fill.color;
        linhaAtual = parseInt(celula.address.split('!')[1].replace(/[A-Z]/g, ''));
        
        //verifica se esá dentro de uma tabela
        if( linhaAtual >= tabelas[indexTabela].linha_ini &&
            linhaAtual < tabelas[indexTabela].linha_fin){
                //se estiver dentro da tabela, valida a cor da célula

                if (celula.color == "#2B2E34"){ //CINZA ESCURO
                    
                    numeracao[i] = [tabelas[indexTabela].index];
                    // console.log(linhaAtual)
                    // console.log(numeracao[i])
                }else if (celula.color == "#FF561C"){ //LARANJA
                    
                    numeracao[i] = ["ITEM"];
                    // console.log(linhaAtual)
                    // console.log(numeracao[i])
                }else if (celula.color == "#FFFFFF" ){ //BRANCO
                    subitem = 1;
                    numeracao[i] = [tabelas[indexTabela].index.toString() + ". " + item ];
                    item ++;
                    // console.log(linhaAtual)
                    // console.log(numeracao[i])
                }else if (celula.color == "#D9D9D9"){ //CINZA
                    numeracao[i] = [tabelas[indexTabela].index.toString() + "." + (item - 1) + "." + subitem];
                    subitem ++
                    // console.log(linhaAtual)
                    // console.log(numeracao[i])
                }else{
                    return -2;
                }
        }else if (linhaAtual == tabelas[indexTabela].linha_fin){
            numeracao[i] = [""];
            indexTabela ++;
            item = 1;
            // console.log(linhaAtual)
            // console.log(numeracao[i])
        }else{ //se não estiver em tabela, preencher com vazio
            numeracao[i] = [""];
            // console.log(linhaAtual)
            // console.log(numeracao[i])
            // console.log(numeracao[i])
        }
    }
    //return numeracao;

    await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem("Precificação");
        var  range = ws.getRange(coluna.concat(primeiraLinha, ':', coluna, ultimaLinha)).load("values");
        //await context.sync();
    
        let novosvalores = numeracao;
        range.values = novosvalores;
        await context.sync();
    });
        
    //se a célula for branca e != da última linha, corresponde a um item
    //se a célular for #D9D9D9, é um subitem

}

async function calculaContribuicao(){
    await atualizaArrayTabelas();

    var arrayContribuicao;
    var custoTotal = 0;
    var linhaFinal = tabelas[tabelas.length -1].linha_fin -1;
    var linhaInicial = tabelas[0].linha_ini + 2;
    //console.log(`custoTotal: ${custoTotal}`)
    var arrayCustos = await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem("Precificação");
        var  range = ws.getRange(colunas[7].fin + linhaInicial + ":" +colunas[7].fin + linhaFinal).load("values");
        //var  range = ws.getRange("AI25").load("values");
        var custoTotal = 0;
        await context.sync();
        return range.values;
    });

    var arrayPropriedades = await getCellProperties(colunas[7].fin + linhaInicial + ":" +colunas[7].fin + linhaFinal);
    //return arrayPropriedades;
    

    //esses blocos FOR são usados para somar somente as células de dentro das tabelas
    // o if mais interno verifica se a cor da fonte da célula é #203764. Caso seja, significa
    //que é o cabeçalho de um kit. Portanto, não irá somar no custo total
    for (j in tabelas){
        for (i in arrayCustos){
            //console.log(`j: ${j}  ---   i: ${i}`)
            //console.log(`if( ${ (parseInt(i) + tabelas[0].linha_ini + 2)}  >=  ${tabelas[j].linha_ini + 2} & ${parseInt(i)+tabelas[0].linha_ini + 2} <= ${tabelas[j].linha_fin -1})`)
            if ( (parseInt(i) + tabelas[0].linha_ini + 2) >= (tabelas[j].linha_ini + 2) & (parseInt(i)+tabelas[0].linha_ini + 2) <= (tabelas[j].linha_fin -1)) {
                if (arrayPropriedades[i][0].format.font.color != "#203764"){
                    custoTotal = custoTotal + parseFloat(arrayCustos[i][0]);
                }else{ //cabeçalho do kit
                    var numSubitens = 0
                    
                    do {
                        numSubitens = numSubitens + 1;
                    } while (arrayPropriedades[parseInt(i) + numSubitens][0].format.fill.color == "#D9D9D9");

                    numSubitens = numSubitens - 1;  
                    //escreve na linha do cabeçalho do kit, coluna de contribuição
                    //a fórmula soma as contribuições dos subitens

                    arrayCustos[i][0] = `=subtotal(9, ${colunas[1].fin + (parseInt(i) + tabelas[0].linha_ini + 3 )  + ":" + colunas[1].fin + (parseInt(i) + tabelas[0].linha_ini + 2 + numSubitens)} )`;
                    //arrayCustos[i][0] = `=${numSubitens}`;
                }
                //console.log(`j: ${j}  ---   i: ${i}`)
                //console.log(`arrayCustos[i][0]: ${arrayCustos[i][0]}`)
                //console.log(`custoTotal: ${custoTotal}`)
            }
            
            if(arrayCustos[i][0] == "VALOR TOTAL"){
                arrayCustos[i][0] = "";
                arrayCustos[i -1 ][0] = "CONTRIBUIÇÃO";
            }
        }
    }

    arrayContribuicao = arrayCustos;

    for (j in tabelas){
        for (i in arrayContribuicao){
            if ( (parseInt(i) + tabelas[0].linha_ini + 2) >= (tabelas[j].linha_ini + 2) & (parseInt(i)+tabelas[0].linha_ini + 2) <= (tabelas[j].linha_fin -1)) {
                if (arrayPropriedades[i][0].format.font.color != "#203764"){
                    arrayContribuicao[i][0] = arred4 (arrayContribuicao[i][0] / custoTotal); 
                    //custoTotal = custoTotal + parseFloat(arrayCustos[i][0]);
                }
            }
        }
    }

    //return arrayContribuicao;
    console.log(custoTotal)
    await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem("Precificação");
        var  range = ws.getRange(colunas[1].fin + linhaInicial + ":" +colunas[1].fin + linhaFinal).load("values");
        
        range.formulas = arrayContribuicao;
        
        await context.sync();
        return arrayContribuicao;
    })
    return arrayContribuicao;
}

async function copiarPlanilhaSV(){

    await Excel.run(async (context) => {
        let myWorkbook = context.workbook;
        let sampleSheet = myWorkbook.worksheets.getItem("{A7441363-1A72-4ACD-854A-C140198E488F}"); //planilha SV
        let precificacao = myWorkbook.worksheets.getItem("{5B74A0A4-C313-D74D-B6C9-894790A73C89}"); //planilha Precificação
        let copiedSheet = sampleSheet.copy("End");
    
        sampleSheet.load("name");
        copiedSheet.load("name");

        precificacao.load("position");
        sampleSheet.load("id");
        
        
        await context.sync();
        copiedSheet.position = precificacao.position + 1;
        copiedSheet.visibility = Excel.SheetVisibility.visible;

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

        await context.sync();
    
        console.log("ID: " + precificacao.id );//+ "' was copied to '" + copiedSheet.name + "'");
    });
}


//recebe um número e retorna arredondado com 2 casas decimais
function arred2 (value) {
    return (Math.round(value*100)/100)
}

//recebe um número e retorna arredondado com 4 casas decimais
function arred4 (value) {
    return (Math.round(value*10000)/10000)
}