function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu("Opcions especials")
  .addItem("Crear pestanya sense_espais", "CrearPestanyaSenseEspais")
  .addItem("Crear pestanya assignatures_uniques", "LlistaAssignatures")
  .addItem("Crear pestanyes assignatures", "CrearPestanyesAssignatures2")
  .addToUi();
}

// Aquesta funció "CrearPestanyaSenseEspais" crea de manera correcta una pestanya 'sense_espais' en la qual
// s'eliminen amb una fórmula els espais generats per un formulari (amb moltes seccions) que genera espais en blanc  

function CrearPestanyaSenseEspais(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const respostesWS = ss.getSheetByName("Respostes al formulari 1");
  
  // Es crea una nova pestanya i se li posa el nom de "sense_espais"
  let ws = ss.insertSheet();
  ws.setName('sense_espais');
  
  // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
  respostesWS.getRange("B1:G1").copyTo(ws.getRange("B1:G1"));
  
  // S'ha pensat aquest codi perquè els headers siguin unics i coincideixin
  ws.getRange("H1").setFormula(`=TRANSPOSE((unique(flatten('Respostes al formulari 1'!H1:W1))))`);
    
  // Es posa Id als HEADER i la formula
  // =SEQUENCE(COUNTA(B2:B))
  ws.getRange("A1").setValue("Id")
  ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)
  
  // Formula d'eliminació d'espais a inserir
  // =ArrayFormula(SUBSTITUTE(SPLIT(TRIM(TRANSPOSE(QUERY(TRANSPOSE(FILTER(SUBSTITUTE('Respostes al formulari 1'!B2:AC;" ";"♥");MMULT(LEN('Respostes al formulari 1'!B2:AC);SEQUENCE(COLUMNS('Respostes al formulari 1'!B2:AC);1;1;0))));;1E+100)));" ");"♥";" "))

  // Insereix la formula a la cel·la B2
  ws.getRange("B2").setFormula(`=ArrayFormula(SUBSTITUTE(SPLIT(TRIM(TRANSPOSE(QUERY(TRANSPOSE(FILTER(SUBSTITUTE('Respostes al formulari 1'!B2:AC;" ";"♥");MMULT(LEN('Respostes al formulari 1'!B2:AC);SEQUENCE(COLUMNS('Respostes al formulari 1'!B2:AC);1;1;0))));;1E+100)));" ");"♥";" "))`)

  // Posem BOLD als Headers
  ws.getRange("A1:W1").setFontWeight("bold");  

}

function LlistaAssignatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const senseEspaisWS = ss.getSheetByName("sense_espais");
  
  // Es crea una nova pestanya i se li posa el nom de "sense_espais"
  const assignaturesWS = ss.insertSheet();
  assignaturesWS.setName('assignatures_uniques');
  
  // Insereix els textes i les fórmules
  // =unique(flatten(sense_espais!H2:M))
  
  
  assignaturesWS.getRange("B1").setValue("Pestanya sense_espais").setFontWeight("bold");
  assignaturesWS.getRange("B2").setValue("Nombre d'alumnes").setFontWeight("bold");
  assignaturesWS.getRange("A3").setFormula(`=unique(flatten(sense_espais!H2:M))`);
  // Posem la fórmula de comptatge a la primera cel·la "B3" que redirigeix a "A3"
  // =ARRAYFORMULA(COUNTIF(sense_espais!H:M;A3))
  assignaturesWS.getRange("B3").setFormula(`=COUNTIF(sense_espais!H:M;A3)`);
  // El Rang B3 el posem en una variable formulaFirstCell
  var formulaFirstCell = assignaturesWS.getRange("B3");
  // Agafem el Rang on copiarem la fórmula i el posem en una variable
  var destinationRange = assignaturesWS.getRange("B3:B26");
  // copiem el contingut de B3 al rang de destinació
  formulaFirstCell.copyTo(destinationRange);
  // Posem la resta de textes en negreta
  assignaturesWS.getRange("A28").setValue("Nombre total d'assignatures escollides").setFontWeight("bold"); ;
  assignaturesWS.getRange("A29").setValue("Alumnes que han escollit").setFontWeight("bold"); ;
  assignaturesWS.getRange("A30").setValue("Nombre d'assignatures que han escollit cada alumne").setFontWeight("bold");
  // Posem la fórmula a la cel·la "B28"
  // =SUM(B3:B26)
  assignaturesWS.getRange("B28").setFormula(`=SUM(B3:B26)`);
   // Posem la fórmula a la cel·la "B28"
  // =B28/B30
  assignaturesWS.getRange("B29").setFormula(`=B28/B30`);
 
  assignaturesWS.autoResizeColumns(1,2);
  assignaturesWS.deleteColumns(3,24);
  assignaturesWS.deleteRows(31,969);
  
 /*
  // De la pestanya `sense_espais`, agafem totes les assignatures del rang que comença a la 2a linia i 8 columna i que es desplaça cap avall fins l'últim i 6 columnes cap a ladreta
  const assignatures = senseEspaisWS
  .getRange(2,8,senseEspaisWS.getLastRow()-1,6)
  .getValues()
     
  console.log(assignatures)
  //// Metode 1
  const grups1 = assignatures.map(nested => nested.map(assignatura => assignatura));
  const assignaturesUniques = [...new Set(...grups1)];
  console.log(grups1)
  
  //// Metode 2
  const grups2 = Array.from(new Set(asignatures.map(assignatura => assignatura)));
  console.log(grups2)
 
  //Posem el mapejat a A3
 
  //assignaturesWS.getRange("A3").setValues(assignaturesUniques)
  
*/  

}

function CrearPestanyesAssignatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const senseEspaisWS = ss.getSheetByName("sense_espais");
  const assignaturesUniquesWS = ss.getSheetByName("assignatures_uniques");

  // De la pestanya `assignatures_uniques`, agafem totes les assignatures del rang que comença a la 2a linia i 8 columna i que es desplaça cap avall fins l'últim i 6 columnes cap a ladreta
  const assignaturesUniques = assignaturesUniquesWS.getRange(3,1,25,1).getValues();
  console.log(assignaturesUniques);
  const assignaturesUniquesArray = Array.from(new Set(assignaturesUniques.map(assignatura => assignatura[0])));
  console.log(assignaturesUniquesArray);
  const assignaturesUniquesArray2 = [...new Set(assignaturesUniques.map(assignatura => assignatura[0]))];
  console.log(assignaturesUniquesArray2);

  // Ens assegurem que no existeixen pestanyes amb els noms ja creades
  
  const currentSheetNames = ss.getSheets().map(s => s.getName());
  
  let ws;

  assignaturesUniquesArray.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE H='Matemàtiques I'";0)
        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE H='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

        // Es posa Id als HEADER i la formula
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id")
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }


  });

}

function CrearPestanyesAssignatures2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const senseEspaisWS = ss.getSheetByName("sense_espais");

  // De la pestanya `sense_espais`, agafem totes les assignatures del rang que comença a la 2a linia i 8 columna i que es desplaça cap avall fins l'últim i 6 columnes cap a ladreta
  const assignatures = senseEspaisWS.getRange(2,8,senseEspaisWS.getLastRow()-1,6).getValues()
  const assignaturesMO = assignatures.map(ass => ass[0]); 
  const assignaturesM1 = assignatures.map(ass => ass[1]);
  const assignaturesM2 = assignatures.map(ass => ass[2]);
  const assignaturesOA1 = assignatures.map(ass => ass[3]);
  const assignaturesOA2 = assignatures.map(ass => ass[4]);
  const assignaturesOA3 = assignatures.map(ass => ass[5]);

  // Deconstruïm
  const assignaturesUniquesMO = [... new Set(assignaturesMO)];
  const assignaturesUniquesM1 = [... new Set(assignaturesM1)];
  const assignaturesUniquesM2 = [... new Set(assignaturesM2)];
  const assignaturesUniquesOA1 = [... new Set(assignaturesOA1)];
  const assignaturesUniquesOA2 = [... new Set(assignaturesOA2)];
  const assignaturesUniquesOA3 = [... new Set(assignaturesOA3)]; 

  // console.log(assignaturesUniquesOA3)

  // Ens assegurem que no existeixen pestanyes amb els noms ja creades
  
  const currentSheetNames = ss.getSheets().map(s => s.getName());
  
  let ws;

  assignaturesUniquesMO.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE H='Matemàtiques I'";0)
        // La lletra H és H perque les assignatures de Modalitat Obligada estan a la columna H
        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE H='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

       // Es posa Id al HEADER "A1" i la formula Id auto a "A2"
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id")
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }


  });


assignaturesUniquesM1.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE I='Matemàtiques I'";0)
        // La lletra I és I perque les assignatures de Modalitat 1 estan a la columna I

        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE I='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

        // Es posa Id al HEADER "A1" i la formula Id auto a "A2"
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id")
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }
});


assignaturesUniquesM2.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE J='Matemàtiques I'";0)
        // La lletra J és J perque les assignatures de Modalitat 2 estan a la columna J
        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE J='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

        // Es posa Id al HEADER "A1" i la formula Id auto a "A2"
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id").setFontWeight("bold");  
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }


  });


assignaturesUniquesOA1.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE K='${assignatura}'";0)`)
        // La lletra K és K perque les assignatures optativa anual 1 estan a la columna K
        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE K='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

        // Es posa Id al HEADER "A1" i la formula Id auto a "A2"
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id").setFontWeight("bold");  
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }
});

assignaturesUniquesOA2.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE L='${assignatura}'";0)`)
        // La lletra L és L perque les assignatures optativa anual 2 estan a la columna L
        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE L='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

        // Es posa Id al HEADER "A1" i la formula Id auto a "A2"
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id").setFontWeight("bold");  
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }
});

assignaturesUniquesOA3.forEach(assignatura => {

    // Comencem bucle If
    // Si els noms de les sheets actuals NO (perquè hi ha el ! davant) inclouen cap assignatura -> La pestanya de l'assignatura no està creada
    if (!currentSheetNames.includes(assignatura)){

        ws = null
        // Insereix una sheet amb el nom de l'assignatura
        ws = ss.insertSheet();
        ws.setName(assignatura)

        // Formula de llistat selectiu a inserir a cada sheet
        // =QUERY(sense_espais!B2:M;"SELECT * WHERE M='${assignatura}'";0)`)
        // La lletra M és M perque les assignatures optativa anual 3 estan a la columna M
        ws.getRange("B2").setFormula(`=QUERY(sense_espais!B2:M;"SELECT * WHERE M='${assignatura}'";0)`)

        // S'agafa el rang dels HEADERS i es copien a la nova pestanya "sense_espais"
        senseEspaisWS.getRange("B1:M1").copyTo(ws.getRange("B1:M1"));

        // Es posa Id al HEADER "A1" i la formula Id auto a "A2"
        // =SEQUENCE(COUNTA(B2:B))
        ws.getRange("A1").setValue("Id").setFontWeight("bold");  
        ws.getRange("A2").setFormula(`=SEQUENCE(COUNTA(B2:B))`)

    }
});

}