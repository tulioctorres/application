//Calculo de distancia DISTÂNCIA

function distance() {
  
  //Inserção dos dados
  var Planilha = "DISTÂNCIA";
  var Start = 1;
  var End = 2;
  var Inicio = 2;
  var colDist = 3;
  
  //Definir planilha e ultima linha
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Planilha);
  var lr = ss.getLastRow();
  
  
  //Inicio do Loop
  for(var i = Inicio; i <= lr; i++){
      
    var EnderStart = ss.getRange(i, Start).getValue();
    var EnderEnd = ss.getRange(i, End).getValue();
    
    var mapObj = Maps.newDirectionFinder();
    mapObj.setOrigin(EnderStart);
    mapObj.setDestination(EnderEnd);
    var directions = mapObj.getDirections();
    
    var km = directions["routes"][0]["legs"][0]["distance"]["value"] * 0.001
    
    ss.getRange(i, colDist).setValue(km);
         
  }  
}
