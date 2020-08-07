"use strict";

var connection = new signalR.HubConnectionBuilder().withUrl("/ImportaDados").build();

connection.on("ReceiveMessage", function (message) {
    var obj = JSON.parse(message);
  
    if (obj.Mensagem == "Fim") {
        alert("Importação concluída.");
    }
    document.getElementById("messagesList").innerHTML = obj.Mensagem;    
});
 

connection.start().then(function () {
    var li = document.createElement("li");
    li.textContent = "Connetado!";
    document.getElementById("messagesList").appendChild(li);
}).catch(function (err) {
    return console.error(err.toString());
});