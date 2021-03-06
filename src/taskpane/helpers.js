
function log(string){
    if (DEBUG){
        console.log(string);
    }
}


async function getFromSheet (sheet, range, load = 'values'){

    return Excel.run(async (context)=>{
      const ws = context.workbook.worksheets.getItem(sheet);
      var  r = ws.getRange(range).load(load);
      
      switch (load) {
        case 'values':
            await context.sync();
            return r.values;
        case 'formulas':
            await context.sync();
            return r.formulas;
        case 'text':
            await context.sync();
            return r.text;
        case 'name':
            ws.load('name');
            await context.sync();
            return ws.name;
        case 'visibility':
            ws.load('visibility');
            await context.sync();
            return ws.visibility;
        default:
            return r.values;
      }
    })
  }
  
  async function writeOnSheet (data, sheet, range, load = 'values'){
    return await Excel.run(async (context)=>{
      const ws = context.workbook.worksheets.getItem(sheet);
      var  r = ws.getRange(range).load(load);
  
      r.values = data;
      switch (load) {
        case 'values':
          r.values = data;
          break;
        case 'formulas':
          r.formulas = data;
          break;
        case 'text':
          r.text = data; 
          break;       
        default:
          r.values = data;
          break;
      }
      return await context.sync();
    });
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
    if(jsdate == '') return ''
    return jsdate.toISOString().split('T')[0];
}

function ExcelDateToJSDate(serialDate, offsetUTC) {
    if (serialDate == '') return '';
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
        log('error');
        log(url);
    }
        
    return obj.value[0].cotacaoVenda;
    //return obj;
    
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
        log("Parâmetro não carregado.");
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

async function funcao(){
    return Excel.run (async(context) =>{
        
        for (let i = tabelas[0].linha_ini+2; i < tabelas[tabelas.length-1].linha_fin ; i ++){
        //for (let i = tabelas[0].linha_ini+2; i < 15 ; i ++){
            let ws = context.workbook.worksheets.getItem(id.precificacao);
            let range = ws.getRange(`M${i}`);
            let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);

            conditionalFormat.custom.rule.formulaLocal = `=SE($R$7=0;0; SE(ABS(M${i} - ARRED((AI${i}/custo_total_projeto);4)) < 0,05%; FALSO; VERDADEIRO))`;
            conditionalFormat.custom.format.fill.color = "#FCE4D6";
        }

        return context.sync();
    })
}

async function novaLinha(){
    const numLinhas = parseInt(document.getElementById("inputNumLinha").value);
    await atualizaArrayTabelas();
    return await Excel.run(async (context) =>{

        var cell = context.workbook.getActiveCell();
        var a = cell.getCellProperties({address: true});
        await context.sync()

        var linhaSelecionada = parseInt( a.value[0][0]["address"].split('!')[1].replace(/[A-Z]/g, '') );
        log(linhaSelecionada);

        if (a.value[0][0]["address"].split('!')[0] != 'Precificação'){
            log('Not in sheet');
            return 0
        }

        const ws = context.workbook.worksheets.getItem(id.precificacao);
        var index_tabela = await estaEmTabela();


        //se fora da tabela
        if (index_tabela == -1){
            log('Not in sheet');
            return "Fora de tabela";
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
        //log(`The address of the used range in the worksheet is "${range.address}"`);
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
    //se a função for chamada, irá excluir a variável cache para excluir linhas, para evitar exclusões acidentais 
    paraExcluir = {
        inicial:-1,
        final:-1
      }
    //pega a ultima linha da aba precificação e salva em variavel
    const ultimaCelula = await getLastCellAddress();
    var celula = {};
    tabelas = [];

    ultimaCelula.coluna = colunas[ colunas["length"] -1 ].fin;

    //copia as propriedades da última coluna para a constante propriedades
    const propriedades = await getCellProperties(ultimaCelula.coluna.concat(1,':', ultimaCelula.coluna, ultimaCelula.linha));
    var numTabelas = 0;
    var indexTabela = 0;
    //log(ultimaCelula.coluna);
    

    var infoTabela = { "index":0 , "num_linhas":0 , "linha_ini":0 , "linha_fin":0 };
    var estaNoKit = false;
    var numLinhasKit = 0;
    var indexKit = 0;

    //varre da primeira até a ultima linha da última coluna
    for(let i = 0; i < ultimaCelula.linha; i++){

        celula = propriedades[i][0];
        //se for uma célula com cor #2B2E34, significa que é uma nova tabela
        if (celula.format.fill.color == "#2B2E34"){

            numTabelas ++;
            tabelas.push({});
            tabelas[numTabelas -1].kit = {};
            tabelas[numTabelas -1].index = numTabelas;
            tabelas[numTabelas -1].linha_ini = i + 1;
            tabelas[numTabelas -1].num_linhas = 0;         
            indexKit = 0;

        } 
        //conta as células internas (com itens) da tabela
        if (!celula.style.includes("Normal") & celula.format.fill.color == "#D9D9D9" & !estaNoKit){
            //log(celula.address);
            //log(celula.style);
            tabelas[numTabelas -1 ].num_linhas ++;
            tabelas[numTabelas -1 ].linha_fin = i + 2;
            tabelas[numTabelas -1 ].kit[indexKit] = {linha: i, subitens: 0};
            estaNoKit = true;

        }

        if (!celula.style.includes("Normal") & celula.format.fill.color == "#D9D9D9" & estaNoKit){
            //log(celula.address);
            //log(celula.style);
            tabelas[numTabelas -1 ].num_linhas ++;
            tabelas[numTabelas -1 ].linha_fin = i + 2;
            tabelas[numTabelas -1 ].kit[indexKit].subitens ++;

            //se a próxima célula não faz parte do kit, incrementa o index
            if(propriedades[i + 1][0].format.fill.color != "#D9D9D9"){
                indexKit ++;
            }
        }
        
        if (!celula.style.includes("Normal") & celula.format.fill.color != "#D9D9D9"){
            //log(celula.address);
            //log(celula.style);
            tabelas[numTabelas -1 ].num_linhas ++;
            tabelas[numTabelas -1 ].linha_fin = i + 2;
            estaNoKit = false;
            
        }

    }


    return tabelas; 
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


  async function novoKit() {
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

  async function estaEmKit(){

    return
  }

  //valida se as linhas do range selecionado estão contidas em uma tabela.
  //se a seleção contiver alguma linha fora da tabela, retorna -1
  //se estiver contido, retorna o index da tabela
  async function rangeNaTabela(){
    await atualizaArrayTabelas();

    return await Excel.run(async (context) => {

        var selecao = await selectedRange();

        if(selecao.planilha != 'Precificação'){
            log('Fora da planilha')
            return -1
        }

        //verifica se o range está contido em alguma das tabelas
        for (i in tabelas){

            //se é a (primeira || segunda || terceira ) && (última || penúltima)
            //se entre terceira e penúltima
            if ((selecao.inicial == tabelas[i].linha_ini || selecao.inicial == tabelas[i].linha_ini + 1 || selecao.inicial == tabelas[i].linha_ini + 2) && selecao.final == tabelas[i].linha_fin){
                log(`está na tabela ${tabelas[i].index}`)
                return tabelas[i].index
                // primeira <= selecao < última
            }else if(selecao.inicial >= tabelas[i].linha_ini && selecao.final < tabelas[i].linha_fin){
                log(`está na tabela ${tabelas[i].index}`)
                return tabelas[i].index
            }else{
                log('Está fora')
            }
        }
        return -1
    })
  }

  async function selectedRange(){
    return await Excel.run(async (context) => {
        var selecao = {
            'inicial' : 0,
            'final': 0,
            'planilha': '',
        }

        var range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();

        selecao.inicial = parseInt(range.address.split('!')[1].split(':')[0].replace(/[A-Z]/g, ''));
        selecao.planilha = range.address.split('!')[0];

        //caso só tenha uma linha selecionada, daria erro no split
        //por isso coloquei o try-catch
        try{
            selecao.final = parseInt(range.address.split('!')[1].split(':')[1].replace(/[A-Z]/g, ''));
        }catch (e){
            selecao.final = parseInt(range.address.split('!')[1].replace(/[A-Z]/g, ''));
        }

        return selecao;
    });
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
        if(listUF[i] == 'GO' ||listUF[i] == 'DF' ||listUF[i] == 'SP'){
            let option = document.createElement("option");
            option.value = listUF[i];
            option.text = listUF[i];

            combo1.appendChild(option);
            //combo2.appendChild(option);
        }
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
    log(`INÍCIO renumerar()`)
    
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
                    // log(linhaAtual)
                    // log(numeracao[i])
                }else if (celula.color == "#FF561C"){ //LARANJA
                    
                    numeracao[i] = ["ITEM"];
                    // log(linhaAtual)
                    // log(numeracao[i])
                }else if (celula.color == "#FFFFFF" ){ //BRANCO
                    subitem = 1;
                    numeracao[i] = [tabelas[indexTabela].index.toString() + ". " + item ];
                    item ++;
                    // log(linhaAtual)
                    // log(numeracao[i])
                }else if (celula.color == "#D9D9D9"){ //CINZA
                    numeracao[i] = [tabelas[indexTabela].index.toString() + "." + (item - 1) + "." + subitem];
                    subitem ++
                    // log(linhaAtual)
                    // log(numeracao[i])
                }else{
                    return -2;
                }
            }else if (linhaAtual == tabelas[indexTabela].linha_fin){
                numeracao[i] = [""];
                indexTabela ++;
            item = 1;
            // log(linhaAtual)
            // log(numeracao[i])
        }else{ //se não estiver em tabela, preencher com vazio
            numeracao[i] = [""];
            // log(linhaAtual)
            // log(numeracao[i])
            // log(numeracao[i])
        }
    }
    //return numeracao;
    
    return await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem("Precificação");
        var  range = ws.getRange(coluna.concat(primeiraLinha, ':', coluna, ultimaLinha)).load("values");
        //await context.sync();
        
        let novosvalores = numeracao;
        range.values = novosvalores;
        log(`FIM renumerar()`)
        return await context.sync();
    });
    
    //se a célula for branca e != da última linha, corresponde a um item
    //se a célular for #D9D9D9, é um subitem
    
}

async function calculaContribuicao(){
    await atualizaArrayTabelas();
    
    var arrayContribuicao;
    var custoTotal = 0;
    var linhaFinal = tabelas[tabelas.length -1].linha_fin;
    var linhaInicial = tabelas[0].linha_ini + 2;
    //log(`custoTotal: ${custoTotal}`)
    var arrayCustos = await Excel.run(async (context)=>{
        const ws = context.workbook.worksheets.getItem(id.precificacao);
        var  range = ws.getRange(colunas[7].fin + linhaInicial + ":" +colunas[7].fin + linhaFinal).load("values");
        //var  range = ws.getRange("AI25").load("values");
        var custoTotal = 0;
        await context.sync();
        return range.values;
    });

    var arrayPropriedades = await getCellProperties(colunas[7].fin + linhaInicial + ":" +colunas[7].fin + linhaFinal);
    log(arrayPropriedades);
    //return arrayPropriedades;
    

    //esses blocos FOR são usados para somar somente as células de dentro das tabelas
    // o if mais interno verifica se a cor da fonte da célula é #203764. Caso seja, significa
    //que é o cabeçalho de um kit. Portanto, não irá somar no custo total
    for (j in tabelas){
        for (i in arrayCustos){
            //log(`j: ${j}  ---   i: ${i}`)
            //log(`if( ${ (parseInt(i) + tabelas[0].linha_ini + 2)}  >=  ${tabelas[j].linha_ini + 2} & ${parseInt(i)+tabelas[0].linha_ini + 2} <= ${tabelas[j].linha_fin -1})`)
            if ( (parseInt(i) + tabelas[0].linha_ini + 2) >= (tabelas[j].linha_ini + 2) & (parseInt(i)+tabelas[0].linha_ini + 2) <= (tabelas[j].linha_fin -1)) {
                if (arrayPropriedades[i][0].format.font.color != "#203764"){
                    custoTotal = custoTotal + parseFloat(arrayCustos[i][0]);
                }else{ //cabeçalho do kit
                    var numSubitens = 0
                    
                    do {
                        numSubitens = numSubitens + 1;

                        log(`arrayPropriedades[parseInt(i) + numSubitens][0].format.fill.color == "#D9D9D9"`)
                        log(`arrayPropriedades[${parseInt(i)} + ${numSubitens}][0].format.fill.color == "#D9D9D9"`)
                        log(parseFloat(arrayCustos[i][0]))
                        log(arrayPropriedades.length)

                    } while ((arrayPropriedades[parseInt(i) + numSubitens][0].format.fill.color == "#D9D9D9") & ((parseInt(i) + numSubitens) < arrayPropriedades.length -1));


                        numSubitens = numSubitens - 1;  
                        //escreve na linha do cabeçalho do kit, coluna de contribuição
                        //a fórmula soma as contribuições dos subitens

                        arrayCustos[i][0] = `=subtotal(9, ${colunas[1].fin + (parseInt(i) + tabelas[0].linha_ini + 3 )  + ":" + colunas[1].fin + (parseInt(i) + tabelas[0].linha_ini + 2 + numSubitens)} )`;
                        //arrayCustos[i][0] = `=${numSubitens}`;
                    

                    
                }
                //log(`j: ${j}  ---   i: ${i}`)
                //log(`arrayCustos[i][0]: ${arrayCustos[i][0]}`)
                //log(`custoTotal: ${custoTotal}`)
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
    log(custoTotal)
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

        if (document.getElementById("input-list-planilha").value == '' ){
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
            copiedSheet.name = document.getElementById("input-list-planilha").value;
            document.getElementById("input-list-planilha").value = "";
        }

        copiedSheet.activate();
        workbook.protection.protect(SECRET)
        copiedSheet.load("id")
        await context.sync();
        id.servicos.push(copiedSheet.id);

        context.workbook.worksheets.getItem(id.param).getRange(COLUNA_ID_SV+ (id.servicos.length + 1) +":"+COLUNA_ID_SV+(id.servicos.length + 1)).values = copiedSheet.id;
    
        log("ID: " + copiedSheet.visibility );//+ "' was copied to '" + copiedSheet.name + "'");

        await atualizaListaPlanilhas()
    });
}

async function novaPlanilhaCustomizada(){

    const COLUNA_ID_CUST = 'Y'

    if(id.custom.length > 10){
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

        if (document.getElementById("input-list-planilha").value == '' ){
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
            novaPlanilha.name = document.getElementById("input-list-planilha").value;
            document.getElementById("input-list-planilha").value = "";
        }

        novaPlanilha.activate();
        workbook.protection.protect(SECRET)
        
        await context.sync();
        id.custom.push(novaPlanilha.id);

        workbook.worksheets.getItem(id.param).getRange(COLUNA_ID_CUST+ (id.custom.length + 1) +":"+COLUNA_ID_CUST+(id.custom.length + 1)).values = novaPlanilha.id;
    
        log("ID: " + novaPlanilha.visibility );//+ "' was copied to '" + copiedSheet.name + "'");

        await atualizaListaPlanilhas()
    });
}

async function cronograma(){
    await atualizaArrayTabelas();

    var coluna_ini = colunas[0].ini;
    var coluna_fin = colunas[0].qtde;
    var headKit = [];
    var numLinhasValidas = tabelas[tabelas.length - 1].linha_fin - tabelas[0].linha_ini - 2 //array com o número de linhas válidas
    
    return await Excel.run(async (context)=>{
        context.workbook.protection.unprotect(SECRET);
        const precificacao = context.workbook.worksheets.getItem("{5B74A0A4-C313-D74D-B6C9-894790A73C89}"); //planilha Precificação
        const cronograma = context.workbook.worksheets.getItem("{4360F843-007A-4860-8658-B6E2AA8612CD}");  //planilha CRONOGRAMA
        cronograma.load("visibility");
        await context.sync();

        var linhaCronograma = cronograma.getRange("3:500");
        var arrayFormulaItem = [["NA","NA","NA","NA","NA","NA","NA","NA","NA"]];
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


        //seleciona as linhas de B até H de todas as linhas com conteúdo
        var origem = precificacao.getRange(
                coluna_ini + (tabelas[0].linha_ini + 2) + ":" +coluna_fin + tabelas[tabelas.length-1].linha_fin);
        
        var offset = (tabelas[0].linha_ini + 2) - 3; //linha original - linha destino

        //copia para a planilha CRONOGRAMA, célula B3
        cronograma.getRange("B3").copyFrom(origem);

        range = "J:J"
        linhaCronograma = cronograma.getRange(range);
        linhaCronograma.numberFormat = "dd/mm/yyyy"

        range = "L:N"
        linhaCronograma = cronograma.getRange(range);
        linhaCronograma.numberFormat = "dd/mm/yyyy"

        range = "P:P"
        linhaCronograma = cronograma.getRange(range);
        linhaCronograma.numberFormat = "dd/mm/yyyy"

        range = "Q:Q"
        linhaCronograma = cronograma.getRange(range);
        linhaCronograma.style = "Currency"

        //escreve as fórmulas dos cronogramas, considerando que não há kits
        for(let i = 3; i < numLinhasValidas + 3; i++){
            range = "I" + (parseInt(i)) + ":" + "Q" + (parseInt(i));
                            //     I    J    K 
            arrayFormulaItem =  [["","","",
                                        `=IF(OR(J${i}="",K${i}=""),"",J${i}+K${i})`, //L
                                        "","","", //M N O
                                        `=if(M${i}=0,"",M${i}+O${i})`, //P
                                        ""]];  //Q
            linhaCronograma = cronograma.getRange(range);
            linhaCronograma.formulas = arrayFormulaItem;
            
            linhaCronograma.format.borders.getItem('EdgeBottom').color = "#000000";
            linhaCronograma.format.borders.getItem('EdgeRight').color = "#000000";
            linhaCronograma.format.borders.getItem('EdgeLeft').color = "#000000";
            linhaCronograma.format.borders.getItem('EdgeTop').color = "#000000";
            //linhaCronograma.format.borders.getItem('InsideHorizontal').color = "#F2F2F2";
            linhaCronograma.format.borders.getItem('InsideVertical').color = "#000000";
            
            range = "O" + (parseInt(i)) + ":" + "O" + (parseInt(i));
            linhaCronograma = cronograma.getRange(range);
            linhaCronograma.format.fill.color = "#FCE4D6"

            range = "Q" + (parseInt(i)) + ":" + "Q" + (parseInt(i));
            linhaCronograma = cronograma.getRange(range);
            linhaCronograma.format.fill.color = "#FCE4D6"

        }
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
            range = "I" + (parseInt(headKit[i].linha - offset)) + ":" + "Q" + (parseInt(headKit[i].linha - offset));
            
            arrayFormulaItem =  [["NA","NA","NA",
                                `=MAX(L${headKit[i].linha - offset +1}:L${headKit[i].linha - offset + headKit[i].subitens})`, //L
                                `=MAX(M${headKit[i].linha - offset +1}:M${headKit[i].linha - offset + headKit[i].subitens})`,  //M
                                "","", // N O
                                `=iferror(M${i}+O${i},"")`, //P
                                ""]];  //Q
            linhaCronograma = cronograma.getRange(range);
            linhaCronograma.formulas = arrayFormulaItem;

            //subitens
            for (let j = 1; j <= headKit[i].subitens; j ++){
                range = "I" + (parseInt(headKit[i].linha - offset + j)) + ":" + "Q" + (parseInt(headKit[i].linha - offset + j));
            
                arrayFormulaItem =  [["","","",  //i J K
                                    `=IF(OR(J${(parseInt(headKit[i].linha - offset + j))}="",K${(parseInt(headKit[i].linha - offset + j))}=""),"",J${(parseInt(headKit[i].linha - offset + j))}+K${(parseInt(headKit[i].linha - offset + j))})`, //L
                                    "", "",  //M N
                                    "NA","NA","NA"]];  //O P Q 
                linhaCronograma = cronograma.getRange(range);
                linhaCronograma.format.borders.getItem('EdgeTop').color = "#F2F2F2";
                if (j < headKit[i].subitens){
                    linhaCronograma.format.borders.getItem('EdgeBottom').color = "#F2F2F2";
                }

                //log(range)
                linhaCronograma.formulas = arrayFormulaItem;

                //linhaCronograma.format.borders.getItem('EdgeBottom').color = "#000000";
                //linhaCronograma.format.borders.getItem('EdgeRight').color = "#000000";
                //linhaCronograma.format.borders.getItem('EdgeLeft').color = "#000000";
                //linhaCronograma.format.borders.getItem('EdgeTop').color = "#F2F2F2";
                //linhaCronograma.format.borders.getItem('InsideHorizontal').color = "#F2F2F2";
                //linhaCronograma.format.borders.getItem('InsideVertical').color = "#F2F2F2";
                
                range = "I" + (parseInt(headKit[i].linha - offset + j)) + ":" + "K" + (parseInt(headKit[i].linha - offset + j));
                linhaCronograma = cronograma.getRange(range);
                linhaCronograma.format.fill.color = "#FCE4D6" //rosa

                range = "M" + (parseInt(headKit[i].linha - offset + j)) + ":" + "N" + (parseInt(headKit[i].linha - offset + j));
                linhaCronograma = cronograma.getRange(range);
                linhaCronograma.format.fill.color = "#FCE4D6" //rosa

                range = "L" + (parseInt(headKit[i].linha - offset + j)) + ":" + "L" + (parseInt(headKit[i].linha - offset + j));
                linhaCronograma = cronograma.getRange(range);
                linhaCronograma.format.fill.color = "#D9D9D9" //cinza

                range = "O" + (parseInt(headKit[i].linha - offset + j)) + ":" + "Q" + (parseInt(headKit[i].linha - offset + j));
                linhaCronograma = cronograma.getRange(range);
                linhaCronograma.format.fill.color = "#D9D9D9" //cinza



            }
          }
        }
                
        //remove as linhas em branco
        for (i in tabelas){
            cronograma.getRange((tabelas[tabelas.length - i -1].linha_fin - offset) + ":" + (tabelas[tabelas.length - i - 1].linha_fin - offset + 3)).delete(Excel.DeleteShiftDirection.up);
        }

        await context.sync();
        cronograma.activate();
        context.workbook.protection.protect(SECRET)
        await context.sync();
        //log("ID: " + cronograma.name );

    });

}

async function copiaTabelaParaDI(){
    log("Início copiaTabelaParaDI()")
    await atualizaArrayTabelas();
    return await Excel.run(async (context)=>{
        const COLOR_BLACK = '#000000' 
        const COLOR_WHITE = '#FFFFFF';
        const CELULA_SOMA_IMPOSTOS = 'D12'
        context.workbook.protection.unprotect(SECRET);
        const primeiraLinha = 21;
        const LINHA_CABECALHO = 20;
        const RANGE_CABECALHO = "B20:G20";
        const precificacao = context.workbook.worksheets.getItem(id.precificacao); 
        const despesas = context.workbook.worksheets.getItem(id.despesas);
        despesas.load("visibility");
        await context.sync();

        //se a planilha já estiver criada, ao pressionar o botão ela será escondida
        if (despesas.visibility == Excel.SheetVisibility.visible){
            despesas.visibility = Excel.SheetVisibility.veryHidden;
            context.workbook.protection.protect(SECRET)
            return context.sync();
        }
        
        despesas.visibility = Excel.SheetVisibility.visible;


        var rangeTabelaOrigem = colunas[0].ini + tabelas[0].linha_ini + ":" + nextLetterInAlphabet(colunas[0].ini, 2) + tabelas[tabelas.length - 1].linha_fin;
        var rangeTabelaOrigem = colunas[0].ini + tabelas[0].linha_ini + ":" + colunas[5].fin + tabelas[tabelas.length - 1].linha_fin;
        var tabelaOrigem = precificacao.getRange(rangeTabelaOrigem).load("values");
        var cabecalhos = [];
       
        context.workbook.protection.protect(SECRET);

        await context.sync()

        despesas.getRange(RANGE_CABECALHO).format.fill.color = COLOR_BLACK;
        despesas.getRange(RANGE_CABECALHO).format.font.color = COLOR_WHITE;
        despesas.getRange(RANGE_CABECALHO).format.horizontalAlignment = 'Center';
        despesas.getRange(RANGE_CABECALHO).format.font.bold = true;
        let header = [['ITEM', 'DESCRIÇÃO', 'VALOR TOTAL', 'TIPO', 'IMPOSTOS (%)', 'IMPOSTOS (R$)']];
        writeOnSheet(header, id.despesas, RANGE_CABECALHO, 'values');
        

        tabelaOrigem = tabelaOrigem.values;

        //filtra somente os cabecalhos e itens da tabela. Desconsidera os subitens
        for ( i in tabelaOrigem){
            if (/^([0-9]+.\s[0-9]+)$/.test(tabelaOrigem[i][0])){
                cabecalhos.push(tabelaOrigem[i])
            }
        }

        //copia itens para aba de despesas
        for (i in cabecalhos){
            despesas.getRange("B" + (parseInt(i) + primeiraLinha)).values = cabecalhos[i][0];
            despesas.getRange("B" + (parseInt(i) + primeiraLinha)).values = cabecalhos[i][0];
            despesas.getRange("E" + (parseInt(i) + primeiraLinha)).values = cabecalhos[i][1].toUpperCase();
            despesas.getRange("C" + (parseInt(i) + primeiraLinha)).values = cabecalhos[i][2];
            despesas.getRange("F" + (parseInt(i) + primeiraLinha)).values = 
                            calculaImpostos(cabecalhos[i][1].toUpperCase(), simNaotoBoolean(cabecalhos[i][23]), simNaotoBoolean(cabecalhos[i][25]));
            despesas.getRange("F" + (parseInt(i) + primeiraLinha)).numberFormat = "0.000%";
            despesas.getRange("G" + (parseInt(i) + primeiraLinha)).formulas = `=D${(parseInt(i) + primeiraLinha)} * F${(parseInt(i) + primeiraLinha)}`

            despesas.getRange("B" + (parseInt(i) + primeiraLinha) + ":" + "G" + (parseInt(i) + primeiraLinha)).format.borders.getItem('EdgeBottom').color = "#000000";
            despesas.getRange("B" + (parseInt(i) + primeiraLinha) + ":" + "G" + (parseInt(i) + primeiraLinha)).format.borders.getItem('EdgeRight').color = "#000000";
            despesas.getRange("B" + (parseInt(i) + primeiraLinha) + ":" + "G" + (parseInt(i) + primeiraLinha)).format.borders.getItem('EdgeLeft').color = "#000000";
            despesas.getRange("B" + (parseInt(i) + primeiraLinha) + ":" + "G" + (parseInt(i) + primeiraLinha)).format.borders.getItem('InsideVertical').color = "#000000";
        }

        //tabela resumo

        
        //await resumo()
        await layoutTableResumo(id.despesas, 'C', 2)
        await contentTableResumo('C', 2)
        
        //soma de impostos
        despesas.getRange(CELULA_SOMA_IMPOSTOS).formulas = `=-SUM(G${primeiraLinha}:G${primeiraLinha + cabecalhos.length - 1})`;
        await context.sync();

        despesas.activate();
        await context.sync();
        return cabecalhos;
        
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

function calculaImpostos(tipoItem, subTrib = false, anexoIX = false){
    log(`Início função calculaImpostos(${tipoItem}, ${subTrib}, ${anexoIX})`)
    var icms = calcICMS(tipoItem, false, subTrib, anexoIX);
    var difal = (param.ufOrig == param.ufDest ? 0: (icmsDaTabela(param.ufDest, param.ufDest) - icms) * !subTrib);

    //log(`icms: ${icms} -- difal: ${difal}`)
    if (param.tipoFatur = "Direto"){
        return trib.fatDireto;
    }

    if (tipoItem == 'HW' || tipoItem == 'MAT'){
        log(`total impostos = csllHW + irpjHW + pis + cofins + icms + difal`);
        log(`total impostos = (${trib.csllHW}) + (${trib.irpjHW}) + (${trib.pis}) + (${trib.cofins}) +(${icms}) +(${difal})`);
        log(`total impostos = ${trib.csllHW + trib.irpjHW + trib.pis + trib.cofins + icms + difal}`);
        log(`Final função calculaImpostos`)
        return trib.csllHW + trib.irpjHW + trib.pis + trib.cofins + icms + difal;
    }else if (tipoItem == 'SW' || tipoItem == 'SV'){
        log(`total impostos = csllSW + irjpSW + pis + cofins + issFora * (!éGoiania) + issGyn * (éGoiania)`)
        log(`total impostos = ${trib.csllSW} + ${trib.irpjSW} + ${trib.pis} + ${trib.cofins} + ${trib.issOut * (!param.destGoiania)} + ${trib.issGYN * (param.destGoiania)}`)
        log(`total impostos = ${trib.csllSW + trib.irpjSW + trib.pis + trib.cofins + trib.issOut * (!param.destGoiania) + trib.issGYN * (param.destGoiania)}`);
        log(`Final função calculaImpostos`)
        return trib.csllSW + trib.irpjSW + trib.pis + trib.cofins + trib.issOut * (!param.destGoiania) + trib.issGYN * (param.destGoiania);
    }else{
        //valida se o tipoItem está dentre os valores permitidos
        log(`Erro: Valor inválido para tipoItem: ${tipoItem}`)
        log(`Final função calculaImpostos`)
        return -1;
    }
}


//recebe um número e retorna arredondado com 2 casas decimais
function arred2 (value) {
    return (Math.round(value*100)/100)
}

//recebe um número e retorna arredondado com 4 casas decimais
function arred4 (value) {
    return (Math.round(value*10000)/10000)
}

function nextLetterInAlphabet(letter, index = 1 ) {
    if (letter == "z") {
      return "a";
    } else if (letter == "Z") {
      return "A";
    } else {
      return String.fromCharCode(letter.charCodeAt(0) + index);
    }
}


//recebe uma string. Se a string estiber escrito SIM/sim, retorna true
//se estiver vazia ou NÃO/não, retorna falso
//1.2 -> adicionei o parâmetro 'vazio'. Caso não seja informado ou FALSE, valores vazios retornam FALSE,
// mas se for passado parâmetro TRUE, valores vazios retornam TRUE
//Obs: não é case sensitive
function simNaotoBoolean (input, vazio = false){
    if(input === ""){
        if (vazio == false){
            return false;
        }else {
            return true;
        }
    }else if(/^([s|S][i|I][m|M])$/.test(input)){
        return true;
    }else if (/^([n|N][a|ã|Ã|A][o|O])$/.test(input)){
        return false;
    }else if( input === true || input === false){
        return input;
    }else{
        log(`Erro em simNaotoBoolean() - input: "${input}"`)
        return -1;
    }
}

async function teste(){
    return await Excel.run(async (context)=>{
        //Excel.createWorkbook(context.workbook);
        //context.workbook.save(Excel.SaveBehavior.prompt);
        const plan = context.workbook.worksheets.getItem("SV")
        //var range = plan.getRange("V2:V12")
        plan.load("id");
        await context.sync();
        //range = range.values;

        //await context.sync()
        //log(plan.id);
        return plan.id
    });

}

//recebe o nome da planilha a ser removida
//remove do excel e do array id
async function removePlanilhaSV(){
    var nome = document.getElementById("input-list-planilha").value;
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
        let lista = document.getElementById("datalist-planilha").options;
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
            let range = param.getRange(COLUNA_ID_CUSTOM+"2"+":"+COLUNA_ID_CUSTOM+"12");
            range.load("values");
            await context.sync();
            for (i in range.values){
                param.getRange(COLUNA_ID_CUSTOM+(2+parseInt(i))).values = ''
                //range.values[i] = ['']
            }

            for (i in id.custom){
                param.getRange(COLUNA_ID_CUSTOM+(2+parseInt(i))).values = id.custom[i]
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
        document.getElementById("input-list-planilha").value = '';
        document.getElementById("btn-add-sheet-sv").disabled = false;
        document.getElementById("btn-add-sheet-br").disabled = false;
        document.getElementById("btn-rem-sheet-sv").disabled = true;
    });
}


async function removeLinha(){

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

async function moveParaDireita(){
    const nome = document.getElementById('input-lista-planilhas').value;
    if (nome == '' ){
        return
    }
   
    if(!existeNaLista(nome)){
        return
    }
    return await Excel.run(async (context)=>{
        const workbook = context.workbook;
        workbook.load("protection/protected");
        var plan = workbook.worksheets.getItem(nome);
        plan.load("position");
        await context.sync();

        if (workbook.protection.protected) {
            workbook.protection.unprotect(SECRET);
        }
        log(plan.position)
        //return plan.position;
        plan.position = plan.position + 1;

        workbook.protection.protect(SECRET);
        await context.sync();
        log(plan.position)
        return context.sync();
        
    });
}

async function moveParaEsquerda(){
    const nome = document.getElementById('input-lista-planilhas').value;
    if (nome == '' ){
        return
    }

    if(!existeNaLista(nome)){
        return
    }
    await Excel.run(async (context)=>{
        const workbook = context.workbook;
        workbook.load("protection/protected");
        var plan = workbook.worksheets.getItem(nome);
        plan.load("position");
        await context.sync();

        if (workbook.protection.protected) {
            workbook.protection.unprotect(SECRET);
        }

        plan.position = plan.position - 1;
        
        workbook.protection.protect(SECRET);
        return context.sync();
    });
}

function inputMargemOnChange(){
    log('inputMargemOnChange()')
    var input_margem                = document.getElementById('content-margem-input-margem');
    var check_comissao_dir_gov      = document.getElementById('content-margem-check-comissao-dir-gov');
    var check_comissao_vp_comercial = document.getElementById('content-margem-check-comissao-vp-comercial');
    var check_comissao_dir_priv     = document.getElementById('content-margem-check-comissao-dir-priv');
    var check_comissao_ger_canais   = document.getElementById('content-margem-check-comissao-ger-canais');
    var check_comissao_exec         = document.getElementById('content-margem-check-comissao-exec');
    var check_comissao_prev         = document.getElementById('content-margem-check-comissao-prev');
    var check_comissao_parc         = document.getElementById('content-margem-check-comissao-parc');
    var input_comissao_dir_gov      = document.getElementById('content-margem-input-comissao-dir-gov');
    var input_comissao_vp_comercial = document.getElementById('content-margem-input-comissao-vp-comercial');
    var input_comissao_dir_priv     = document.getElementById('content-margem-input-comissao-dir-priv');
    var input_comissao_ger_canais   = document.getElementById('content-margem-input-comissao-ger-canais');
    var input_comissao_exec         = document.getElementById('content-margem-input-comissao-exec');
    var input_comissao_prev         = document.getElementById('content-margem-input-comissao-prev');
    var input_comissao_parc         = document.getElementById('content-margem-input-comissao-parc');
    var input_tx_admin              = document.getElementById('content-margem-input-tx-admin');
    var input_sv_terc               = document.getElementById('content-margem-input-sv-terc');
    var check_proj_estrategico      = document.getElementById('content-margem-check-proj-estrategico');
    var check_politica_automatica   = document.getElementById('content-margem-check-politica-automatica');


    if (!check_politica_automatica.checked){
        log('check_politica_automatica = false')
        loadFromMargemToSheet();
        return 0
    }

    if (check_comissao_dir_gov.checked){
        let range = getRangeMargem(input_margem.value);
        input_comissao_dir_gov.value = comissoes.diretoria_governo[range] * 100;
    }

    if (check_comissao_dir_priv.checked){
        let range = getRangeMargem(input_margem.value);
        input_comissao_dir_priv.value = comissoes.diretoria_privado[range] * 100;
    }
    
    if (check_comissao_vp_comercial.checked){
        let range = getRangeMargem(input_margem.value);
        input_comissao_vp_comercial.value = comissoes.vp_comercial[range] * 100;
    }

    if (check_comissao_ger_canais.checked){
        let range = getRangeMargem(input_margem.value);
        input_comissao_ger_canais.value = comissoes.gerente_canais[range] * 100;
    }

    if (check_comissao_exec.checked){
        let range = getRangeMargem(input_margem.value);
        input_comissao_exec.value = comissoes.executivo[range] * 100;
    }

    if (check_comissao_parc.checked){
        let range = getRangeMargem(input_margem.value);
        input_comissao_parc.value = comissoes.parceiro[range] * 100;
    }

    if (check_comissao_prev.checked){
        (check_proj_estrategico.checked) ? 
        input_comissao_prev.value = comissoes.prevendas['projeto_estrategico'] * 100 :
        input_comissao_prev.value = comissoes.prevendas['projeto_comum'] * 100;
    }
    loadFromMargemToSheet();

}

function margemComissoesOnChange(){
    var input_margem                = document.getElementById('content-margem-input-margem');
    var check_comissao_dir_gov      = document.getElementById('content-margem-check-comissao-dir-gov');
    var check_comissao_vp_comercial = document.getElementById('content-margem-check-comissao-vp-comercial');
    var check_comissao_dir_priv     = document.getElementById('content-margem-check-comissao-dir-priv');
    var check_comissao_ger_canais   = document.getElementById('content-margem-check-comissao-ger-canais');
    var check_comissao_exec         = document.getElementById('content-margem-check-comissao-exec');
    var check_comissao_prev         = document.getElementById('content-margem-check-comissao-prev');
    var check_comissao_parc         = document.getElementById('content-margem-check-comissao-parc');
    var check_proj_estrategico      = document.getElementById('content-margem-check-proj-estrategico');
    var check_politica_automatica   = document.getElementById('content-margem-check-politica-automatica');

    var input_comissao_dir_gov      = document.getElementById('content-margem-input-comissao-dir-gov');
    var input_comissao_vp_comercial = document.getElementById('content-margem-input-comissao-vp-comercial');
    var input_comissao_dir_priv     = document.getElementById('content-margem-input-comissao-dir-priv');
    var input_comissao_ger_canais   = document.getElementById('content-margem-input-comissao-ger-canais');
    var input_comissao_exec         = document.getElementById('content-margem-input-comissao-exec');
    var input_comissao_prev         = document.getElementById('content-margem-input-comissao-prev');
    var input_comissao_parc         = document.getElementById('content-margem-input-comissao-parc');
    var range = getRangeMargem(input_margem.value);

    if (check_comissao_dir_gov.checked){
            if (check_politica_automatica.checked) input_comissao_dir_gov.value = comissoes.diretoria_governo[range] * 100;
    } else  input_comissao_dir_gov.value = 0;

    if (check_comissao_dir_priv.checked){
        if (check_politica_automatica.checked) input_comissao_dir_priv.value = comissoes.diretoria_privado[range] * 100;
    } else  input_comissao_dir_priv.value = 0;
    
    if (check_comissao_vp_comercial.checked){
        if (check_politica_automatica.checked) input_comissao_vp_comercial.value = comissoes.vp_comercial[range] * 100;
    } else  input_comissao_vp_comercial.value = 0;

    if (check_comissao_ger_canais.checked){
        if (check_politica_automatica.checked) input_comissao_ger_canais.value = comissoes.gerente_canais[range] * 100;
    } else  input_comissao_ger_canais.value = 0;

    if (check_comissao_exec.checked){
        if (check_politica_automatica.checked) input_comissao_exec.value = comissoes.executivo[range] * 100;
    } else  input_comissao_exec.value = 0;

    if (check_comissao_parc.checked){
        if (check_politica_automatica.checked) input_comissao_parc.value = comissoes.parceiro[range] * 100;
    } else  input_comissao_parc.value = 0;

    if (check_comissao_prev.checked){
        if (check_politica_automatica.checked) {
            (check_proj_estrategico.checked) ? 
            input_comissao_prev.value = comissoes.prevendas['projeto_estrategico'] * 100 :
            input_comissao_prev.value = comissoes.prevendas['projeto_comum'] * 100;
        }
    } else {
        input_comissao_prev.value = 0;
    }

    loadFromMargemToSheet();
}


function getRangeMargem(x) {
    if (x <= 0) return 'equal_0';
    if (0 < x && x < 10) return 'btw_0_10';
    if (x == 10) return 'equal_10';
    if (10 < x && x < 15) return 'btw_10_15';
    if (x == 15) return 'equal_15';
    if (15 < x && x < 20) return 'btw_15_20';
    if (x == 20) return 'equal_20';
    if (20 < x && x < 25) return 'btw_20_25';
    if (x == 25) return 'equal_25';
    if (25 < x && x < 30) return 'btw_25_30';
    if (x == 30) return 'equal_30';
    if (30 < x && x < 35) return 'btw_30_35';
    if (x == 35) return 'equal_35';
    if (35 < x && x < 40) return 'btw_35_40';
    if (x == 40) return 'equal_40';
    if (40 < x && x < 45) return 'btw_40_45';
    if (x == 45) return 'equal_45';
    if (45 < x && x < 50) return 'btw_45_50';
    if (x >= 50) return 'equal_gr_50';
}


//o objeto input date do frontend retorna value no formato YYYY-MM-DD
//essa função converte para formato EUA: MM-DD-YYY, que é aceito pela API do BC
function dataHTMLparaEUA(dataHTML){
    dataHTML = dataHTML.split('-');
    dataHTML.push(dataHTML[0]);
    dataHTML.shift();
    return dataHTML.join('-');
}

//retorna uma string no formato YYYY-MM-DD
//o valor da data retornado é (dataFinal - intervalo)
function calculaDataInicial (dataFinal, intervalo){
    let df = dataFinal.split('-').map(Number);
    let date = new Date(df[0], df[1] - 1, df[2]);
    date.setDate(date.getDate() - intervalo);
    return date.toISOString().split('T')[0];
}

//essa função busca a cotação da moeda em um determinado dia e em um intervalo de dias atrás.
//ela retorna a cotação mais próxima da data passada como parâmetro.
//o intervalo é necessário pois em dias não comerciais, não há cotação. Para evitar que
//a busca retorne vazia, o intervalo garante que haverá cotações válidas.
async function ptaxQuotation(moeda, dataFinal = ontemStringHTML(), intervalo = 5){
  
    let obj = null;
    var dataInicial = calculaDataInicial (dataFinal, intervalo);
    dataInicial = dataHTMLparaEUA(dataInicial);
    dataFinal = dataHTMLparaEUA(dataFinal);
    let url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?@moeda='" + moeda + "'&@dataInicial='" + dataInicial + "'&@dataFinalCotacao='" + dataFinal + "'&$top=100&$filter=tipoBoletim%20eq%20'Fechamento'&$orderby=dataHoraCotacao%20desc&$format=json&$select=cotacaoVenda,dataHoraCotacao,tipoBoletim";
  
    try {
        obj = await (await fetch(url)).json();
    } catch(e) {
        log('error');
        log(url);
    }
        
    return obj.value[0].cotacaoVenda;
    //return obj;   
  }

  
async function saveParamToSheet(){
    const range_param_values = 'B1:B29';
    var array_param = [];
    
    array_param = 
    [
      [param.tipoFatur],
      [param.margem],
      [param.comissaoDirGov],
      [param.comissaoExec],
      [param.comissaoPrev],
      [param.comissaoParc],
      [param.txAdm],
      [param.svTerc],
      [param.aplicarComissaoDirGov],
      [param.aplicarComissaoExec],
      [param.aplicarComissaoPrev],
      [param.aplicarComissaoParc],
      [param.ufOrig],
      [param.ufDest],
      [param.destGoiania],
      [param.zerarICMS],
      [param.txImpHW],
      [param.txImpSW],
      [JSDateToStrDate(param.dataCambio)],
      [param.dolarPTAX],
      [param.euroPTAX],
      [param.politicaAutom],
      [param.comissaoDirPriv],
      [param.comissaoVP],
      [param.comissaoGC],
      [param.aplicarComissaoDirPriv],
      [param.aplicarComissaoVP],
      [param.aplicarComissaoGC],
      [param.projetoEstrategico]
    ];
  
    await writeOnSheet(array_param, id.param, range_param_values, 'values')
  }

//GV -> frontend (menu parametros)
async function writeParametrosOnFrontend(){

    var input_tx_import_hw = document.getElementById('content-parametros-input-tx-hw')
    var input_tx_import_sw = document.getElementById('content-parametros-input-tx-sw')
  
    input_tx_import_hw.value = param.txImpHW;
    input_tx_import_sw.value = param.txImpSW;
}
  
//planilha -> frontend (edição - nomes das planilhas customizadas)
async function writeSheetsNamesOnFrontend(){
    var datalist_planilhas = document.getElementById('datalist-planilhas');
    datalist_planilhas.innerHTML = '';

    for (i in id.servicos){
        if (i != 0) {
        let option = document.createElement('option');
        option.value = await getFromSheet(id.servicos[i], '', 'name');
        option.text = 'Planilha de Serviços';
        datalist_planilhas.appendChild(option);
        }
    }

    for (i in id.custom){
        let option = document.createElement('option');
        option.value = await getFromSheet(id.custom[i], '', 'name');
        option.text = 'Planilha Customizada';
        datalist_planilhas.appendChild(option);
    }
}

//GV -> frontend (cambio e faturamento)
function writeCambioOnFrontend(){
var data_cambio = document.getElementById('content-cambio-input-data')
var input_usd = document.getElementById('content-cambio-input-usd')
var input_eur = document.getElementById('content-cambio-input-eur')
var select_uf_orig  = document.getElementById('content-cambio-select-uf-orig')
var select_uf_dest = document.getElementById('content-cambio-select-uf-dest')
var select_tipo_faturamento  = document.getElementById('content-cambio-select-tipo-fat')
var check_incentivo_icms = document.getElementById('content-cambio-check-icms')
var select_destino_goiania = document.getElementById('content-cambio-dest-gyn')

//carrega lista de UF Destinos
for (i in listUF) {
    let option = document.createElement("option");
    option.value = listUF[i];
    option.text = listUF[i];
    select_uf_dest.appendChild(option);
}

//carrega lista de UF Origens
for (i in listUF) {
    if(listUF[i] == 'GO' ||listUF[i] == 'DF' ||listUF[i] == 'SP'){
        let option = document.createElement("option");
        option.value = listUF[i];
        option.text = listUF[i];
        select_uf_orig.appendChild(option);
    }
}

select_uf_orig.value = param.ufOrig;
select_uf_dest.value = param.ufDest;
data_cambio.value = JSDateToStrDate(param.dataCambio);
data_cambio.max = ontemStringHTML();
input_usd.value = param.dolarPTAX;
input_eur.value = param.euroPTAX;
select_tipo_faturamento.value = param.tipoFatur;
check_incentivo_icms.checked = param.zerarICMS;
select_destino_goiania.checked = param.destGoiania;
}

//GV -> frontend (margem e comissoes)
function writeMargemOnFrontend(){
var input_margem                = document.getElementById('content-margem-input-margem');
var check_comissao_dir_gov      = document.getElementById('content-margem-check-comissao-dir-gov');
var check_comissao_vp_comercial = document.getElementById('content-margem-check-comissao-vp-comercial');
var check_comissao_dir_priv     = document.getElementById('content-margem-check-comissao-dir-priv');
var check_comissao_ger_canais   = document.getElementById('content-margem-check-comissao-ger-canais');
var check_comissao_exec         = document.getElementById('content-margem-check-comissao-exec');
var check_comissao_prev         = document.getElementById('content-margem-check-comissao-prev');
var check_comissao_parc         = document.getElementById('content-margem-check-comissao-parc');
var input_comissao_dir_gov      = document.getElementById('content-margem-input-comissao-dir-gov');
var input_comissao_vp_comercial = document.getElementById('content-margem-input-comissao-vp-comercial');
var input_comissao_dir_priv     = document.getElementById('content-margem-input-comissao-dir-priv');
var input_comissao_ger_canais   = document.getElementById('content-margem-input-comissao-ger-canais');
var input_comissao_exec         = document.getElementById('content-margem-input-comissao-exec');
var input_comissao_prev         = document.getElementById('content-margem-input-comissao-prev');
var input_comissao_parc         = document.getElementById('content-margem-input-comissao-parc');
var input_tx_admin              = document.getElementById('content-margem-input-tx-admin');
var input_sv_terc               = document.getElementById('content-margem-input-sv-terc');
var check_proj_estrategico      = document.getElementById('content-margem-check-proj-estrategico');
var check_politica_automatica   = document.getElementById('content-margem-check-politica-automatica');


input_margem.value = param.margem * 100;
input_comissao_dir_gov.value = param.comissaoDirGov * 100;
input_comissao_vp_comercial.value = param.comissaoVP * 100;
input_comissao_dir_priv.value = param.comissaoDirPriv * 100;
input_comissao_ger_canais.value = param.comissaoGC * 100;
input_comissao_exec.value = param.comissaoExec * 100;
input_comissao_prev.value = param.comissaoPrev * 100;
input_comissao_parc.value = param.comissaoParc * 100;
input_tx_admin.value = param.txAdm * 100;
input_sv_terc.value = param.svTerc * 100;
check_comissao_dir_gov.checked = param.aplicarComissaoDirGov;
check_comissao_vp_comercial.checked = param.aplicarComissaoVP;
check_comissao_dir_priv .checked = param.aplicarComissaoDirPriv;
check_comissao_ger_canais.checked = param.aplicarComissaoGC;
check_comissao_exec.checked = param.aplicarComissaoExec;
check_comissao_prev.checked = param.aplicarComissaoPrev;
check_comissao_parc.checked = param.aplicarComissaoParc;
check_proj_estrategico.checked = param.projetoEstrategico;
check_politica_automatica.checked = param.politicaAutom;
}

//frontend (menu edicao)-> GV -> planilha
function loadFromEdicaoToSheet(){
    return 0
  }
  
//frontend (menu cambio e faturamento)-> GV -> planilha
async function loadFromCambioToSheet(){
var data_cambio               = document.getElementById('content-cambio-input-data')
var input_usd                 = document.getElementById('content-cambio-input-usd')
var input_eur                 = document.getElementById('content-cambio-input-eur')
var select_uf_orig            = document.getElementById('content-cambio-select-uf-orig')
var select_uf_dest            = document.getElementById('content-cambio-select-uf-dest')
var select_tipo_faturamento   = document.getElementById('content-cambio-select-tipo-fat')
var check_incentivo_icms      = document.getElementById('content-cambio-check-icms')
var select_destino_goiania    = document.getElementById('content-cambio-dest-gyn')

data_cambio.value == '' ? param.dataCambio = '': param.dataCambio = new Date(data_cambio.value + 'T00:00:00');
param.dolarPTAX                 = input_usd.value;
param.euroPTAX                  = input_eur.value;
param.ufOrig                    = select_uf_orig.value;
param.ufDest                    = select_uf_dest.value;
param.tipoFatur                 = select_tipo_faturamento.value;
param.zerarICMS                 = check_incentivo_icms.checked;
param.destGoiania               = select_destino_goiania.checked;

await saveParamToSheet();

}

//frontend (menu margem e comissoes)-> GV -> planilha
async function loadFromMargemToSheet(){
var input_margem                = document.getElementById('content-margem-input-margem');
var check_comissao_dir_gov      = document.getElementById('content-margem-check-comissao-dir-gov');
var check_comissao_vp_comercial = document.getElementById('content-margem-check-comissao-vp-comercial');
var check_comissao_dir_priv     = document.getElementById('content-margem-check-comissao-dir-priv');
var check_comissao_ger_canais   = document.getElementById('content-margem-check-comissao-ger-canais');
var check_comissao_exec         = document.getElementById('content-margem-check-comissao-exec');
var check_comissao_prev         = document.getElementById('content-margem-check-comissao-prev');
var check_comissao_parc         = document.getElementById('content-margem-check-comissao-parc');
var input_comissao_dir_gov      = document.getElementById('content-margem-input-comissao-dir-gov');
var input_comissao_vp_comercial = document.getElementById('content-margem-input-comissao-vp-comercial');
var input_comissao_dir_priv     = document.getElementById('content-margem-input-comissao-dir-priv');
var input_comissao_ger_canais   = document.getElementById('content-margem-input-comissao-ger-canais');
var input_comissao_exec         = document.getElementById('content-margem-input-comissao-exec');
var input_comissao_prev         = document.getElementById('content-margem-input-comissao-prev');
var input_comissao_parc         = document.getElementById('content-margem-input-comissao-parc');
var input_tx_admin              = document.getElementById('content-margem-input-tx-admin');
var input_sv_terc               = document.getElementById('content-margem-input-sv-terc');
var check_proj_estrategico      = document.getElementById('content-margem-check-proj-estrategico');
var check_politica_automatica   = document.getElementById('content-margem-check-politica-automatica');

param.margem                  = arred4(input_margem.value / 100);
param.comissaoDirGov          = arred4(input_comissao_dir_gov.value / 100);
param.comissaoVP              = arred4(input_comissao_vp_comercial.value / 100);
param.comissaoDirPriv         = arred4(input_comissao_dir_priv.value / 100);
param.comissaoGC              = arred4(input_comissao_ger_canais.value / 100);
param.comissaoExec            = arred4(input_comissao_exec.value / 100);
param.comissaoPrev            = arred4(input_comissao_prev.value / 100);
param.comissaoParc            = arred4(input_comissao_parc.value / 100);
param.txAdm                   = arred4(input_tx_admin.value / 100);
param.svTerc                  = arred4(input_sv_terc.value / 100);
param.aplicarComissaoDirGov   = check_comissao_dir_gov.checked;
param.aplicarComissaoVP       = check_comissao_vp_comercial.checked;
param.aplicarComissaoDirPriv  = check_comissao_dir_priv .checked;
param.aplicarComissaoGC       = check_comissao_ger_canais.checked;
param.aplicarComissaoExec     = check_comissao_exec.checked;
param.aplicarComissaoPrev     = check_comissao_prev.checked;
param.aplicarComissaoParc     = check_comissao_parc.checked;
param.projetoEstrategico      = check_proj_estrategico.checked;
param.politicaAutom           = check_politica_automatica.checked;

saveParamToSheet();
}

//frontend (menu parametros)-> GV -> planilha
async function loadFromParametrosToSheet(){

var input_tx_import_hw = document.getElementById('content-parametros-input-tx-hw')
var input_tx_import_sw = document.getElementById('content-parametros-input-tx-sw')

param.txImpHW = input_tx_import_hw.value;
param.txImpSW = input_tx_import_sw.value;

saveParamToSheet();

}

  async function loadFromParamSheetV1(){
    const range_param_values = 'B1:B29';
    const range_id_sv_sheet = 'V3:V12';
    const range_id_custom_sheet = 'Y3:Y12';
    //const range_policy = 'F2:K24';
  
    var param_values = await getFromSheet(id.param, range_param_values, 'values')
  
    param.hoje                    = new Date();
    param.tipoFatur               = param_values[0][0];
    param.margem                  = param_values[1][0];
    param.comissaoDirGov          = param_values[2][0];
    param.comissaoExec            = param_values[3][0];
    param.comissaoPrev            = param_values[4][0];
    param.comissaoParc            = param_values[5][0];
    param.txAdm                   = param_values[6][0];
    param.svTerc                  = param_values[7][0];
    param.aplicarComissaoDirGov   = param_values[8][0];
    param.aplicarComissaoExec     = param_values[9][0];
    param.aplicarComissaoPrev     = param_values[10][0];
    param.aplicarComissaoParc     = param_values[11][0];
    param.ufOrig                  = param_values[12][0];
    param.ufDest                  = param_values[13][0];
    param.destGoiania             = param_values[14][0];
    param.zerarICMS               = param_values[15][0];
    param.txImpHW                 = param_values[16][0];
    param.txImpSW                 = param_values[17][0];
    param.dataCambio              = ExcelDateToJSDate(param_values[18][0], -3);
    param.dolarPTAX               = param_values[19][0];
    param.euroPTAX                = param_values[20][0];
    param.politicaAutom           = param_values[21][0];
    param.comissaoDirPriv         = param_values[22][0];
    param.comissaoVP              = param_values[23][0];
    param.comissaoGC              = param_values[24][0];
    param.aplicarComissaoDirPriv  = param_values[25][0];
    param.aplicarComissaoVP       = param_values[26][0];
    param.aplicarComissaoGC       = param_values[27][0];
    param.projetoEstrategico      = param_values[28][0];
    
  
    var id_sv_sheet = await getFromSheet(id.param, range_id_sv_sheet, 'values')
    
    for (i in id_sv_sheet){
      if (id_sv_sheet[i] != ''){
        id.servicos.push(id_sv_sheet[i][0]);
      }
    }
    
    var id_custom_sheet = await getFromSheet(id.param, range_id_custom_sheet, 'values')
  
    for (i in id_custom_sheet){
      if (id_custom_sheet[i] != ''){
        id.custom.push(id_custom_sheet[i][0]);
      }
    }
  }
  
  async function loadFromParamSheetV2(){
    const range_param_values = 'B1:B29';
    const range_id_sv_sheet = 'V3:V12';
    const range_id_custom_sheet = 'Y3:Y12';
    const range_policy = 'F2:K24';
  
    var param_values = await getFromSheet(id.param, range_param_values, 'values')
  
    param.hoje                    = new Date();
    param.tipoFatur               = param_values[0][0];
    param.margem                  = param_values[1][0];
    param.comissaoDirGov          = param_values[2][0];
    param.comissaoExec            = param_values[3][0];
    param.comissaoPrev            = param_values[4][0];
    param.comissaoParc            = param_values[5][0];
    param.txAdm                   = param_values[6][0];
    param.svTerc                  = param_values[7][0];
    param.aplicarComissaoDirGov   = param_values[8][0];
    param.aplicarComissaoExec     = param_values[9][0];
    param.aplicarComissaoPrev     = param_values[10][0];
    param.aplicarComissaoParc     = param_values[11][0];
    param.ufOrig                  = param_values[12][0];
    param.ufDest                  = param_values[13][0];
    param.destGoiania             = param_values[14][0];
    param.zerarICMS               = param_values[15][0];
    param.txImpHW                 = param_values[16][0];
    param.txImpSW                 = param_values[17][0];
    param.dataCambio              = ExcelDateToJSDate(param_values[18][0], -3);
    param.dolarPTAX               = param_values[19][0];
    param.euroPTAX                = param_values[20][0];
    param.politicaAutom           = param_values[21][0];
    param.comissaoDirPriv         = param_values[22][0];
    param.comissaoVP              = param_values[23][0];
    param.comissaoGC              = param_values[24][0];
    param.aplicarComissaoDirPriv  = param_values[25][0];
    param.aplicarComissaoVP       = param_values[26][0];
    param.aplicarComissaoGC       = param_values[27][0];
    param.projetoEstrategico      = param_values[28][0];
  
    var policy_values = await getFromSheet(id.param, range_policy, 'values');
  
    comissoes.diretoria_governo.equal_0     = policy_values[0][0];
    comissoes.diretoria_governo.btw_0_10    = policy_values[1][0];
    comissoes.diretoria_governo.equal_10    = policy_values[2][0];
    comissoes.diretoria_governo.btw_10_15   = policy_values[3][0];
    comissoes.diretoria_governo.equal_15    = policy_values[4][0];
    comissoes.diretoria_governo.btw_15_20   = policy_values[5][0];
    comissoes.diretoria_governo.equal_20    = policy_values[6][0];
    comissoes.diretoria_governo.btw_20_25   = policy_values[7][0];
    comissoes.diretoria_governo.equal_25    = policy_values[8][0];
    comissoes.diretoria_governo.btw_25_30   = policy_values[9][0];
    comissoes.diretoria_governo.equal_30    = policy_values[10][0];
    comissoes.diretoria_governo.btw_30_35   = policy_values[11][0];
    comissoes.diretoria_governo.equal_35    = policy_values[12][0];
    comissoes.diretoria_governo.btw_35_40   = policy_values[13][0];
    comissoes.diretoria_governo.equal_40    = policy_values[14][0];
    comissoes.diretoria_governo.btw_40_45   = policy_values[15][0];
    comissoes.diretoria_governo.equal_45    = policy_values[16][0];
    comissoes.diretoria_governo.btw_45_50   = policy_values[17][0];
    comissoes.diretoria_governo.equal_gr_50 = policy_values[18][0];
  
    comissoes.diretoria_privado.equal_0     = policy_values[0][1];
    comissoes.diretoria_privado.btw_0_10    = policy_values[1][1];
    comissoes.diretoria_privado.equal_10    = policy_values[2][1];
    comissoes.diretoria_privado.btw_10_15   = policy_values[3][1];
    comissoes.diretoria_privado.equal_15    = policy_values[4][1];
    comissoes.diretoria_privado.btw_15_20   = policy_values[5][1];
    comissoes.diretoria_privado.equal_20    = policy_values[6][1];
    comissoes.diretoria_privado.btw_20_25   = policy_values[7][1];
    comissoes.diretoria_privado.equal_25    = policy_values[8][1];
    comissoes.diretoria_privado.btw_25_30   = policy_values[9][1];
    comissoes.diretoria_privado.equal_30    = policy_values[10][1];
    comissoes.diretoria_privado.btw_30_35   = policy_values[11][1];
    comissoes.diretoria_privado.equal_35    = policy_values[12][1];
    comissoes.diretoria_privado.btw_35_40   = policy_values[13][1];
    comissoes.diretoria_privado.equal_40    = policy_values[14][1];
    comissoes.diretoria_privado.btw_40_45   = policy_values[15][1];
    comissoes.diretoria_privado.equal_45    = policy_values[16][1];
    comissoes.diretoria_privado.btw_45_50   = policy_values[17][1];
    comissoes.diretoria_privado.equal_gr_50 = policy_values[18][1];
  
    comissoes.vp_comercial.equal_0     = policy_values[0][2];
    comissoes.vp_comercial.btw_0_10    = policy_values[1][2];
    comissoes.vp_comercial.equal_10    = policy_values[2][2];
    comissoes.vp_comercial.btw_10_15   = policy_values[3][2];
    comissoes.vp_comercial.equal_15    = policy_values[4][2];
    comissoes.vp_comercial.btw_15_20   = policy_values[5][2];
    comissoes.vp_comercial.equal_20    = policy_values[6][2];
    comissoes.vp_comercial.btw_20_25   = policy_values[7][2];
    comissoes.vp_comercial.equal_25    = policy_values[8][2];
    comissoes.vp_comercial.btw_25_30   = policy_values[9][2];
    comissoes.vp_comercial.equal_30    = policy_values[10][2];
    comissoes.vp_comercial.btw_30_35   = policy_values[11][2];
    comissoes.vp_comercial.equal_35    = policy_values[12][2];
    comissoes.vp_comercial.btw_35_40   = policy_values[13][2];
    comissoes.vp_comercial.equal_40    = policy_values[14][2];
    comissoes.vp_comercial.btw_40_45   = policy_values[15][2];
    comissoes.vp_comercial.equal_45    = policy_values[16][2];
    comissoes.vp_comercial.btw_45_50   = policy_values[17][2];
    comissoes.vp_comercial.equal_gr_50 = policy_values[18][2];
  
    comissoes.gerente_canais.equal_0     = policy_values[0][3];
    comissoes.gerente_canais.btw_0_10    = policy_values[1][3];
    comissoes.gerente_canais.equal_10    = policy_values[2][3];
    comissoes.gerente_canais.btw_10_15   = policy_values[3][3];
    comissoes.gerente_canais.equal_15    = policy_values[4][3];
    comissoes.gerente_canais.btw_15_20   = policy_values[5][3];
    comissoes.gerente_canais.equal_20    = policy_values[6][3];
    comissoes.gerente_canais.btw_20_25   = policy_values[7][3];
    comissoes.gerente_canais.equal_25    = policy_values[8][3];
    comissoes.gerente_canais.btw_25_30   = policy_values[9][3];
    comissoes.gerente_canais.equal_30    = policy_values[10][3];
    comissoes.gerente_canais.btw_30_35   = policy_values[11][3];
    comissoes.gerente_canais.equal_35    = policy_values[12][3];
    comissoes.gerente_canais.btw_35_40   = policy_values[13][3];
    comissoes.gerente_canais.equal_40    = policy_values[14][3];
    comissoes.gerente_canais.btw_40_45   = policy_values[15][3];
    comissoes.gerente_canais.equal_45    = policy_values[16][3];
    comissoes.gerente_canais.btw_45_50   = policy_values[17][3];
    comissoes.gerente_canais.equal_gr_50 = policy_values[18][3];
  
    comissoes.executivo.equal_0     = policy_values[0][4];
    comissoes.executivo.btw_0_10    = policy_values[1][4];
    comissoes.executivo.equal_10    = policy_values[2][4];
    comissoes.executivo.btw_10_15   = policy_values[3][4];
    comissoes.executivo.equal_15    = policy_values[4][4];
    comissoes.executivo.btw_15_20   = policy_values[5][4];
    comissoes.executivo.equal_20    = policy_values[6][4];
    comissoes.executivo.btw_20_25   = policy_values[7][4];
    comissoes.executivo.equal_25    = policy_values[8][4];
    comissoes.executivo.btw_25_30   = policy_values[9][4];
    comissoes.executivo.equal_30    = policy_values[10][4];
    comissoes.executivo.btw_30_35   = policy_values[11][4];
    comissoes.executivo.equal_35    = policy_values[12][4];
    comissoes.executivo.btw_35_40   = policy_values[13][4];
    comissoes.executivo.equal_40    = policy_values[14][4];
    comissoes.executivo.btw_40_45   = policy_values[15][4];
    comissoes.executivo.equal_45    = policy_values[16][4];
    comissoes.executivo.btw_45_50   = policy_values[17][4];
    comissoes.executivo.equal_gr_50 = policy_values[18][4];
  
    comissoes.parceiro.equal_0     = policy_values[0][5];
    comissoes.parceiro.btw_0_10    = policy_values[1][5];
    comissoes.parceiro.equal_10    = policy_values[2][5];
    comissoes.parceiro.btw_10_15   = policy_values[3][5];
    comissoes.parceiro.equal_15    = policy_values[4][5];
    comissoes.parceiro.btw_15_20   = policy_values[5][5];
    comissoes.parceiro.equal_20    = policy_values[6][5];
    comissoes.parceiro.btw_20_25   = policy_values[7][5];
    comissoes.parceiro.equal_25    = policy_values[8][5];
    comissoes.parceiro.btw_25_30   = policy_values[9][5];
    comissoes.parceiro.equal_30    = policy_values[10][5];
    comissoes.parceiro.btw_30_35   = policy_values[11][5];
    comissoes.parceiro.equal_35    = policy_values[12][5];
    comissoes.parceiro.btw_35_40   = policy_values[13][5];
    comissoes.parceiro.equal_40    = policy_values[14][5];
    comissoes.parceiro.btw_40_45   = policy_values[15][5];
    comissoes.parceiro.equal_45    = policy_values[16][5];
    comissoes.parceiro.btw_45_50   = policy_values[17][5];
    comissoes.parceiro.equal_gr_50 = policy_values[18][5];
  
    comissoes.prevendas.projeto_comum       = policy_values[21][0];
    comissoes.prevendas.projeto_estrategico = policy_values[22][0];
  
    var id_sv_sheet = await getFromSheet(id.param, range_id_sv_sheet, 'values')
    
    for (i in id_sv_sheet){
      if (id_sv_sheet[i] != ''){
        id.servicos.push(id_sv_sheet[i][0]);
      }
    }
    
    var id_custom_sheet = await getFromSheet(id.param, range_id_custom_sheet, 'values')
  
    for (i in id_custom_sheet){
      if (id_custom_sheet[i] != ''){
        id.custom.push(id_custom_sheet[i][0]);
      }
    }
  }

  async function loadFromTribSheetV1(){
    const range_trib_values = 'B2:B29';
    const range_tabela_icms_values = 'F3:AF29';
  
    var trib_values = await getFromSheet(id.trib, range_trib_values, 'values');
  
    trib.irpjHW       = trib_values[0][0]
    trib.irpjSW       = trib_values[1][0]
    trib.csllHW       = trib_values[4][0]
    trib.csllSW       = trib_values[5][0]
    trib.cppHW        = trib_values[8][0]
    trib.cppSW        = trib_values[9][0]
    trib.issGYN       = trib_values[12][0]
    trib.issOut       = trib_values[13][0]
    trib.pis          = trib_values[16][0]
    trib.cofins       = trib_values[20][0]
  
    //valida se há informação na celula trib!B29. Se não tiver, escreve 11%.
    //em versões mais antigas da FPV, esse campo não existia
    if (trib_values[27][0] == ''){
      trib.fatDireto = 0.11
      await writeOnSheet(0.11, id.trib, 'B29', 'values')
    } else{
      trib.fatDireto = trib_values[27][0];
    }
  
    var tabela_icms_values = await getFromSheet(id.trib, range_tabela_icms_values, 'values')
  
    trib.tabelaIcms = tabela_icms_values;
  
  }

  async function loadFromTribSheetV2(){
      await loadFromTribSheetV1();
  }
  