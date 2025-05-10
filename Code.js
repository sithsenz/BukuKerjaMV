const UI = SpreadsheetApp.getUi();
const lembaran = SpreadsheetApp.getActiveSheet();


function onOpen() {
  UI.createMenu('Verifikasi')
    .addSubMenu(
      UI.createMenu('Semak Data')
        .addItem('Varians', 'paparSemakData')
    )

    .addSubMenu(
      UI.createMenu('Perbandingan Kaedah')
        .addItem('OLS', 'bandingOLS')
        .addItem('Deming', 'bandingODeming')
    )

    .addItem('Linearity', 'semakLinearity')

    .addToUi();
}


function paparSemakData() {
  let barTepi = HtmlService.createHtmlOutputFromFile('semakData');
  barTepi.setTitle('Semak Data');
  UI.showSidebar(barTepi);
}


function paparPilihX() {
  let kawasanX = lembaran.getActiveRange().getA1Notation();

  return [kawasanX, 'kawasanX'];
}


function paparPilihY() {
  let kawasanY = lembaran.getActiveRange().getA1Notation();

  return [kawasanY, 'kawasanY'];
}


function semakVarians(data) {
  lembaran.getRange('A1').setFormula(`=FTEST(${data[0]},${data[1]})`).setNumberFormat('0.0000');

  let pFvarians = lembaran.getRange('A1').getDisplayValue();
  let kesimpulanFvarians = pFvarians < 0.05? 'Varians berubah-ubah' : 'Varians seragam';

  lembaran.getRange('A1').setValue('#');

  return [kesimpulanFvarians, pFvarians];
}