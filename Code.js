/*========
1. Pemalar
========*/
const UI = SpreadsheetApp.getUi();
const lembaran = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Utama');
const bukuKerja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BukuKerja');


/*=============
2. Menu Tersuai
=============*/
function onOpen() {
  UI.createMenu('Verifikasi')
    .addItem('Kira Purata SD', 'paparKiraPurataSD')
    .addItem('Perbandingan Kaedah', 'paparBandingan')
    .addItem('Linearity', 'semakLinearity')
    .addToUi();
}


/*==================
3. Fungsi Menu Utama
==================*/
function paparKiraPurataSD() {
  let menuTepi = HtmlService.createHtmlOutputFromFile('kiraPurataSD');
  menuTepi.setTitle('Kira Purata dan SD Data');
  UI.showSidebar(menuTepi);
}

function paparBandingan() {
  let menuTepi = HtmlService.createHtmlOutputFromFile('bandingan');
  menuTepi.setTitle('Perbandingan Kaedah');
  UI.showSidebar(menuTepi);
}


/*==============================
4. Fungsi Bantuan Kira Purata SD
==============================*/
function kiraPurataSDgs(kawasan) {
  let dataY = lembaran.getRange(kawasan[1]).getValues();
  let bilSampel = dataY.length;

  if (kawasan[0] != "") {
    let dataX = lembaran.getRange(kawasan[0]).getValues();

    for (let i=0; i<bilSampel; i++) {
      setPurataSD(lembaran, i, 8, dataX);
    }
  }

  for (let i=0; i<bilSampel; i++) {
    setPurataSD(lembaran, i, 10, dataY);
  }

  let kawasanPurataSD = lembaran.getRange(2,8,bilSampel,4);
  kawasanPurataSD.setValues(kawasanPurataSD.getValues());
  kawasanPurataSD.setNumberFormat('0.0000');
}


/*===========================
5. Fungsi Perbandingan Kaedah
===========================*/
function kiraWLSRgs(kawasan) {
  let dataX = lembaran.getRange(`Utama!${kawasan[0]}`).getValues();
  let dataY = lembaran.getRange(`Utama!${kawasan[1]}`).getValues();
  let dataSDy = lembaran.getRange(`Utama!${kawasan[2]}`).getValues();
  let bilSampel = dataX.length;

  let dataBeratY = [];
  let dataBeratX = [];
  let dataPintasan = [];

  for (let i=0; i<bilSampel; i++) {
    let pemberat = dataSDy[i][0];
    dataBeratY.push([dataY[i][0] / pemberat]);
    dataBeratX.push([dataX[i][0] / pemberat]);
    dataPintasan.push([1 / pemberat]);
  }

  bukuKerja.getRange(1,1,bilSampel,1).setValues(dataBeratY);
  bukuKerja.getRange(1,2,bilSampel,1).setValues(dataBeratX);
  bukuKerja.getRange(1,3,bilSampel,1).setValues(dataPintasan);

  let kawasanY = bukuKerja.getRange(1,1,bilSampel,1).getA1Notation();
  let kawasanXsdY = bukuKerja.getRange(1,2,bilSampel,2).getA1Notation();
  bukuKerja.getRange('E1').setFormula(`=LINEST(${kawasanY},${kawasanXsdY},FALSE,TRUE)`);

  let dataCoef = bukuKerja.getRange('E1').getDataRegion().getValues();

  bukuKerja.getRange('E8').setFormula(`=TINV(0.05,F4)`);
  let nilaiStudentT = bukuKerja.getRange('E8').getValue();
  let SEMbeta0 = nilaiStudentT * dataCoef[1][0];
  let SEMbeta1 = nilaiStudentT * dataCoef[1][1];

  bukuKerja.getRange('E10').setValue(dataCoef[0][0] - SEMbeta0);
  bukuKerja.getRange('F10').setValue(dataCoef[0][0]);
  bukuKerja.getRange('G10').setValue(dataCoef[0][0] + SEMbeta0);
  bukuKerja.getRange('E11').setValue(dataCoef[0][1] - SEMbeta1);
  bukuKerja.getRange('F11').setValue(dataCoef[0][1]);
  bukuKerja.getRange('G11').setValue(dataCoef[0][1] + SEMbeta1);
  bukuKerja.getRange('E10').getDataRegion().setNumberFormat('0.0000')

  bukuKerja.getRange('E13').setValue(dataCoef[2][0]).setNumberFormat('0.00%');

  let beta0 = bukuKerja.getRange('F10').getDisplayValue();
  let beta0bawah = bukuKerja.getRange('E10').getDisplayValue();
  let beta0atas = bukuKerja.getRange('G10').getDisplayValue();

  let beta1 = bukuKerja.getRange('F11').getDisplayValue();
  let beta1bawah = bukuKerja.getRange('E11').getDisplayValue();
  let beta1atas = bukuKerja.getRange('G11').getDisplayValue();

  let rKuasa2 = `R2 = ${bukuKerja.getRange('E13').getDisplayValue()}`;

  let persamaan = `Y = (${beta0}) + (${beta1})X`;
  let kecerunan = `Kecerunan = (${beta1bawah} , ${beta1atas})`;
  let pintasan = `Pintasan = (${beta0bawah} , ${beta0atas})`;

  return [persamaan, kecerunan, pintasan, rKuasa2];
}


/*====================
6. Fungsi Bantuan Umum
====================*/
function paparPilih(x) {
  let kawasanTerpilih = {
    'x': 'kawasanX',
    'sdx': 'kawasanSDx',
    'y': 'kawasanY',
    'sdy': 'kawasanSDy',
  }
  let kawasan = lembaran.getActiveRange().getA1Notation();

  return [kawasan, kawasanTerpilih[x]];
}

function setPurataSD(lembar, baris, lajurPurata, data) {
  lembar.getRange(2+baris, lajurPurata).setFormula(`=AVERAGE(${data[baris]})`);
  lembar.getRange(2+baris, lajurPurata+1).setFormula(`=IFERROR(STDEV(${data[baris]}), "")`);
}