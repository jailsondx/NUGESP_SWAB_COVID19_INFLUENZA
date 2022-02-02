function hora_mod() {

    var today = new Date();
    var dia = String(today.getDate()).padStart(2, "0");
    var mes = String(today.getMonth() + 1).padStart(2, "0"); //January is 0!
    var ano = today.getFullYear();
    var hora = String(today.getHours() + 2).padStart(2, "0");
    var min = String(today.getMinutes()).padStart(2, "0");
    var sec = String(today.getSeconds()).padStart(2, "0");

    today = dia + "/" + mes + "/" + ano + " às "+hora+":"+min+":"+sec;
    //today = Utilities.formatDate(new Date(today), "GMT-300", "dd/MM/yyyy hh:mm:ss aa");
    Logger.log("Hoje: " + today);

    return today;
}
