var base = "C:/Users/Administrador/Documents/Sistemas/PDFgenerator/indesign";

var TEMPLATE_PRODUTOS = File(base + "/template/template_produtos.indd");
var TEMPLATE_CONTRA = File(base + "/template/template_contracapa.indd");

var CSVPRODUTOS = File(base + "/CSV/data_merge_produto.csv");
var CSVCONTRACAPA = File(base + "/CSV/data_merge_contracapa.csv");

var OUTFOLDER = Folder(base + "/output");
var PDF_PRODUTOS = File(base + "/output/_produtos.pdf");
var PDF_CONTRA = File(base + "/output/_contracapa.pdf");
var PDF_FINAL = File(base + "/output/resultado.pdf");

if (!OUTFOLDER.exists) OUTFOLDER.create();

function gerarPDF(template, csv, pdfOut) {
    var doc = app.open(template);
    doc.dataMergeProperties.selectDataSource(csv);  
    doc.dataMergeProperties.updateDataSource();
    doc.dataMergeProperties.mergeRecords();
    var merged = app.activeDocument;
    merged.exportFile(ExportFormat.pdfType, pdfOut);
    merged.close(SaveOptions.no);
    doc.close(SaveOptions.no);
}

var oldUI = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

try {
    gerarPDF(TEMPLATE_PRODUTOS, CSVPRODUTOS, PDF_PRODUTOS);
    gerarPDF(TEMPLATE_CONTRA, CSVCONTRACAPA, PDF_CONTRA);

    var finalDoc = app.documents.add();

    var produtos = app.pdfPlacePreferences;

    var pdfProdDoc = app.open(PDF_PRODUTOS);
    var totalPaginasProdutos = pdfProdDoc.pages.length;
    pdfProdDoc.close(SaveOptions.no);

    produtos.pageNumber = 1;
    finalDoc.pages[0].place(PDF_PRODUTOS);

    for (var i = 2; i <= totalPaginasProdutos; i++) {
        produtos.pageNumber = i;
        var novaPage = finalDoc.pages.add();
        novaPage.place(PDF_PRODUTOS);
    }

    var page = finalDoc.pages.add();
    page.place(PDF_CONTRA);

    finalDoc.exportFile(ExportFormat.pdfType, PDF_FINAL);
    finalDoc.close(SaveOptions.no);

} catch (e) {
    alert("❌ Erro ao gerar PDF:/n" + e);
}

app.scriptPreferences.userInteractionLevel = oldUI;
