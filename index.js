// Importation de la librairie XLSX pour travailler avec les fichiers Excel
const XLSX = require('xlsx');
//instanciation de la librairie n-readlines permettant lectures par ligne d'un fichier d'entrée.
const lineByLine = require('n-readlines');

// Lecture du fichier en entrée.
const liner = new lineByLine('input.csv');

//variable de lecture d'une ligne lue
var line;
//le numéro de ligne
var lineNumber = 0;
//tableau de valeur
var tableauGTP = [];
var objetGTP;
var lineASCII;

let _Matricule;let _CodeElt; let _CodeRubrique; let _Valeur;

while (line = liner.next()) {
    //convertion de la ligne en mode ASCII
    lineASCII = line.toString('ascii');

    if(lineASCII.startsWith('VM')){
        _Matricule = lineASCII.slice(2,15).trim();
        _CodeElt=0;
        _CodeRubrique = lineASCII.slice(15,32).trim();
        _Valeur = lineASCII.slice(32).trim()

        objetGTP = [_Matricule,_CodeElt,_CodeRubrique,_Valeur];
        tableauGTP.push(objetGTP);

        //aller à la ligne suivante
        lineNumber++;
    }
}

console.log(tableauGTP);

// Creation d'un nouveau workbook Excel
var workbook = XLSX.utils.book_new();
//Creation d'un feuille excel
var worksheet = XLSX.utils.aoa_to_sheet(tableauGTP);
//Ajout de la feuille au workbook
XLSX.utils.book_append_sheet(workbook, worksheet, "resultats");
// Noms d'en-têtes du fichier
XLSX.utils.sheet_add_aoa(worksheet, [["Matricule", "Code Elt", "Code Rubrique", "Valeur"]], { origin: "A1" });
/* create an XLSX file and try to save to resultats.xlsx */
XLSX.writeFile(workbook, "resultats.xlsx");
