/*========
1. Pemalar
========*/
const UI = SpreadsheetApp.getUi();
const lembaran = SpreadsheetApp.getActiveSheet();


/*=============
2. Menu Tersuai
=============*/
function onOpen() {
  UI.createMenu('Verifikasi')
    .addItem('Semak Data', 'paparSemakData')

    .addSubMenu(
      UI.createMenu('Perbandingan Kaedah')
        .addItem('OLS', 'bandingOLS')
        .addItem('Deming', 'bandingODeming')
    )

    .addItem('Linearity', 'semakLinearity')

    .addToUi();
}


/*==================
3. Fungsi Menu Utama
==================*/
function paparSemakData() {
  let menuTepi = HtmlService.createHtmlOutputFromFile('semakData');
  menuTepi.setTitle('Semak Data');
  UI.showSidebar(menuTepi);
}


/*==========================
4. Fungsi Bantuan Semak Data
==========================*/
function paparPilih(x) {
  let kawasanTerpilih = x == 'x'? 'kawasanX' : 'kawasanY';
  let kawasan = lembaran.getActiveRange().getA1Notation();

  return [kawasan, kawasanTerpilih];
}

function semakDataGS(data) {
  let selA1 = lembaran.getRange('A1');
  const dataAsalA1 = selA1.getDisplayValue();

  let selB1 = lembaran.getRange('B1');
  const dataAsalB1 = selB1.getDisplayValue();

  let selC1 = lembaran.getRange('C1');
  const dataAsalC1 = selC1.getDisplayValue();

  selA1.setFormula(`=FTEST(${data[0]},${data[1]})`).setNumberFormat('0.0000');

  let pFvarians = selA1.getDisplayValue();
  let kesimpulanFvarians = pFvarians < 0.05? 'Varians berubah-ubah' : 'Varians seragam';

  selB1.setFormula(`=QUARTILE(${data[1]},1)`);
  selC1.setFormula(`=QUARTILE(${data[1]},3)`);

  let julatKuartil = selC1.getValue() - selB1.getValue();
  let hadBawah = selB1.getValue() - 1.5 * julatKuartil;
  let hadAtas = selC1.getValue() + 1.5 * julatKuartil;

  selB1.setValue(hadBawah).setNumberFormat('0.0000');
  selC1.setValue(hadAtas).setNumberFormat('0.0000');
  let kesimpulanNormal = `${selB1.getDisplayValue()} <= y <= ${selC1.getDisplayValue()}`;

  const julatDiformat = lembaran.getRange(data[1]);
  const formatBersyarat = SpreadsheetApp.newConditionalFormatRule()
                          .whenNumberNotBetween(hadBawah, hadAtas)
                          .setBackground('#FF0000')
                          .setRanges([julatDiformat])
                          .build();

  const peraturanFormatBersyarat = [];
  peraturanFormatBersyarat.push(formatBersyarat);
  lembaran.setConditionalFormatRules(peraturanFormatBersyarat);

  selA1.setValue(dataAsalA1);
  selB1.setValue(dataAsalB1);
  selC1.setValue(dataAsalC1);

  return [kesimpulanFvarians, pFvarians, kesimpulanNormal];
}