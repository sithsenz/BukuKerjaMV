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


/*============================================
5. Fungsi Perbandingan Kaedah - WLS Regression
============================================*/
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

  let beta0 = pemformat().format(dataCoef[0][0]);
  let beta0bawah = pemformat().format(dataCoef[0][0] - SEMbeta0);
  let beta0atas = pemformat().format(dataCoef[0][0] + SEMbeta0);

  let beta1 = pemformat().format(dataCoef[0][1]);
  let beta1bawah = pemformat().format(dataCoef[0][1] - SEMbeta1);
  let beta1atas = pemformat().format(dataCoef[0][1] + SEMbeta1);

  let rKuasa2 = `R2 = ${pemformat('percent', 2).format(dataCoef[2][0])}`;

  let persamaan = `Y = (${beta0}) + (${beta1})X`;
  let kecerunan = `Kecerunan = [${beta1bawah} , ${beta1atas}]`;
  let pintasan = `Pintasan = [${beta0bawah} , ${beta0atas}]`;

  return [persamaan, kecerunan, pintasan, rKuasa2];
}


/*=====================================================
Fungsi Perbandingan Kaedah - Weighted Deming Regression
=====================================================*/
function wDemingR(kawasan) {
  const purata_x = [];
  const purata_y = [];
  const sd_x = [];
  const sd_y = [];
  const lam = [];

  let dataX = lembaran.getRange(`Utama!${kawasan[0]}`).getValues();
  let dataY = lembaran.getRange(`Utama!${kawasan[1]}`).getValues();
  let bilSampel = dataX.length;

  dataX.forEach(baris => {purata_x.push(bantuKiraPurata(baris))});
  dataY.forEach(baris => {purata_y.push(bantuKiraPurata(baris))});

  dataX.forEach(baris => {sd_x.push(bantuKiraSDSampel(baris))});
  dataY.forEach(baris => {sd_y.push(bantuKiraSDSampel(baris))});

  for (let i=0; i<bilSampel; i++) {
    lam.push((sd_x[i]**2) / (sd_y[i]**2));
  }

  let beta1 = 1.0;
  let beta0 = 0.0;
  let kejituan = 0.00001;
  let bilMaxPercubaan = 100;
  let bilPercubaan = 0;
  let tercapai = false;
  let pemberat = [];

  while (!tercapai && bilPercubaan < bilMaxPercubaan) {
    pemberat = purata_x.map((_, i) => {
      let berat = sd_x[i]**2 + (beta1 * sd_y[i])**2;
      return 1 / berat;
    });

    const stats = bantuKiraWLS(purata_x, purata_y, pemberat);
    const beta1Baru = stats.kecerunan;
    const beta0Baru = stats.pintasan;

    if (Math.abs(beta1Baru - beta1) < kejituan) {
      tercapai = true;
    }

    beta1 = beta1Baru;
    beta0 = beta0Baru;
    bilPercubaan++;
  }

  const {r2, sse} = bantuKiraR2(purata_x, purata_y, pemberat, beta1, beta0);
  const selangKeyakinan = bantuKiraCI(purata_x, purata_y, pemberat, beta1, beta0, bilSampel, sse);

  let rKuasa2 = `R2 = ${pemformat('percent', 2).format(r2)}`;

  let persamaan = `Y = (${pemformat().format(beta0)}) + (${pemformat().format(beta1)})X`;
  let kecerunan = `Kecerunan = [${pemformat().format(selangKeyakinan.cerunBawah)} , ${pemformat().format(selangKeyakinan.cerunAtas)}]`;
  let pintasan = `Pintasan = [${pemformat().format(selangKeyakinan.pintasBawah)} , ${pemformat().format(selangKeyakinan.pintasAtas)}]`;

  return [persamaan, kecerunan, pintasan, rKuasa2];
}

function bantuKiraR2(x, y, berat, cerun, pintas) {
  let sse = 0;
  let sst = 0;

  const purataBeratY = bantuKiraPurata(y.map((yi, i) => berat[i] * yi)) / bantuKiraPurata(berat);

  for (let i=0; i<x.length; i++) {
    const b = berat[i];
    const bakiResidu = y[i] - (pintas + cerun * x[i]);
    sse += b * bakiResidu**2;
    sst += b * (y[i] - purataBeratY)**2
  }

  const r2 = 1 - (sse / sst);
  return {r2, sse};
}

function bantuKiraCI(x, y, berat, cerun, pintas, bil, sse) {
  let jumBerat = 0;
  let jumBeratX = 0;
  let jumBeratX2 = 0;

  for (let i=0; i<bil; i++) {
    const b = berat[i];
    jumBerat += b;
    jumBeratX += b * x[i];
    jumBeratX2 += b * x[i]**2;
  }

  const mse = sse / (bil - 2);
  const seCerun = Math.sqrt(mse / (jumBeratX2 - (jumBeratX**2 / jumBerat)));
  const sePintas = Math.sqrt(mse * (1 / jumBerat + jumBeratX**2 / (jumBerat**2 * jumBeratX2)));

  bukuKerja.getRange('E8').setFormula(`=TINV(0.05,${bil - 2})`);
  const nilaiStudentT = bukuKerja.getRange('E8').getValue();

  return {
    cerunBawah: cerun - nilaiStudentT * seCerun,
    cerunAtas: cerun + nilaiStudentT * seCerun,
    pintasBawah: pintas - nilaiStudentT * sePintas,
    pintasAtas: pintas + nilaiStudentT * sePintas,
  };
}

function bantuKiraWLS(x, y, berat) {
  let jumBerat = 0;
  let jumBeratX = 0;
  let jumBeratY = 0;
  let jumBeratXY = 0;
  let jumBeratX2 = 0;

  for (let i=0; i<x.length; i++) {
    const b = berat[i];
    jumBerat += b;
    jumBeratX += b * x[i];
    jumBeratY += b * y[i];
    jumBeratXY += b * x[i] * y[i]
    jumBeratX2 += b * x[i]**2;
  }

  const kecerunan = (jumBerat * jumBeratXY - jumBeratX * jumBeratY) / (jumBerat * jumBeratX2 - jumBeratX**2);
  const pintasan = (jumBeratY - kecerunan * jumBeratX) / jumBerat;

  return {kecerunan, pintasan};
}

function bantuKiraPurata(data) {
  let purata = (data.reduce((a, b) => a + b, 0)) / data.length;
  return purata;
}

function bantuKiraSDSampel(data) {
  let p = bantuKiraPurata(data);
  let v = (data.reduce((a, b) => a + (b - p)**2, 0)) / (data.length - 1);
  let s = v > 0? Math.sqrt(v) : 1e-6;
  return s;
}


/*====================
7. Fungsi Bantuan Umum
====================*/
function pemformat(stail='decimal', bilTp=4) {
  const pemf = new Intl.NumberFormat('en-US', {
    style: stail,
    minimumFractionDigits: bilTp,
  });

  return pemf;
}


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