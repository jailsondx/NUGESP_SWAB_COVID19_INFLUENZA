/* 
  
  ANOTAÇÕES
  
  NAVEGUE PELA PLANILHA SE GUIANDO COMO UMA MATRIZ:...
  VAR.FUNC(PARAMENTROS)[LINHA][COLUNA];
  .getDataRange() 0 x 0;
  .getRange() 1x1;
  
  ASPAS DUPLAS PARA TEXTO.
  SIMBOLO DE SOMA(+) PARA CONTATENAR.

  .getFullYear(); (que retorna 2018), em vez de getYear(), que retorna o valor do ano indexado em 1900 (ou seja, neste caso, retornaria 118).
  .padStart() para completar com zero à esquerda os valores menores que 10 (assim, 5 é mostrado como 05). Não funciona no IE
  
*/
function Solicitacao_Testes(sheet) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RESULTADOS');
    var planilha = sheet.getDataRange();
    var linha = 5;
    var ultima_linha = sheet.getRange("A5").getDataRegion().getLastRow();
    var ultimo_enviado = sheet.getRange("V5").getDataRegion().getLastRow();
    var email_assunto = "AVISO DE SOLICITAÇÃO DE RT-PCR COVID-19";
    var email_coordenador;
    var nome_colaborador;
    var cargo;
    var setor;
    var data_solicitacao
    var nome_coordenador;
    var covid;
    var influenza;
    var checkbox;
    var emailTemp = HtmlService.createTemplateFromFile("EmailSolicitação");
    var htmlMessage;

    //Executa o loop de envio de email e definição de STATUS, com ponto de parada como a ultima linha NOME preenchida
    for (linha = 539; linha < ultima_linha; linha++) {
        covid = planilha.getValues()[linha][13];
        influenza = planilha.getValues()[linha][14];
        checkbox = planilha.getValues()[linha][20];
        if (checkbox == false) {
            Logger.log("CHECKBOX: " + checkbox);
            if (((covid == "") || (covid == null)) && ((influenza == "") || (influenza == null))) {
                Logger.log("COVID: " + covid);
                Logger.log("INFLUENZA: " + influenza);
                nome_colaborador = planilha.getValues()[linha][0];
                cargo = planilha.getValues()[linha][3];
                setor = planilha.getValues()[linha][2];
                data_solicitacao = planilha.getValues()[linha][6];
                nome_coordenador = planilha.getValues()[linha][7];
                email_coordenador = planilha.getValues()[linha][8];
                Email2(linha, sheet, email_assunto, email_coordenador, nome_coordenador, nome_colaborador, cargo, setor, data_solicitacao, emailTemp, htmlMessage);
                //DefineStatus2(linha, sheet);
                Logger.log("Colaborador: " + nome_colaborador + " || Coordenador: " + nome_coordenador);
            } else {}
        } else {}
    } //FIM FOR
    Logger.log("Ultima linha com dados: " + ultima_linha);
    Logger.log("Ultima linha enviada: " + ultimo_enviado);

    //Alert();

} //FIM DA FUNÇÃO SOLICITAÇÃO_TESTE

//Função que preenche a coluna Status
function DefineStatus2(linha, sheet) {
    var status = sheet.getRange(linha + 1, 21); //Define o intervalo STATUS
    status.setValue(true); //Seta valor ENVIADO no intervalo definido
}

function Email2(linha, sheet, email_assunto, email_coordenador, nome_coordenador, nome_colaborador, cargo, setor, data_solicitacao, emailTemp, htmlMessage) {

    var format_data_solicitacao;
    format_data_solicitacao = Utilities.formatDate(new Date(data_solicitacao), "GMT -300", "dd/MM/yyyy");

    //ENVIO DE EMAIL
    emailTemp.htmlnomecoordenador = nome_coordenador;
    emailTemp.htmlnomecolaborador = nome_colaborador;
    emailTemp.htmlcargo = cargo;
    emailTemp.htmlsetor = setor;
    emailTemp.htmlsolicitacao = format_data_solicitacao;
    htmlMessage = emailTemp.evaluate().getContent();
    GmailApp.sendEmail(
        email_coordenador,
        email_assunto,
        "You email does't support HTML.", {
            htmlBody: htmlMessage
        });
    DefineStatus2(linha, sheet);

} //FECHA FUNÇÃO EMAIL


//Função de Alerta
function Alert2() {
    var interface = SpreadsheetApp.getUi(); //variavel de interface
    interface.alert("SCRIPT FINALIZADO, SOLICITAÇÕES NOTIFICADAS")
    /*
    if (cont > 0){
      interface.alert("SUCESSO! ENVIADO(S) " + cont + " NOVO(S) EMAIL(S)"); //modulo de alerta
    } else {
      interface.alert("NENHUM NOVO EMAIL ENVIADO"); //modulo de alerta
    }
    */
}
