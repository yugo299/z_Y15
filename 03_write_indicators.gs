function functionI() {

  if (pSheet.getRange(tRow, 1).getValue()) { return }

  let row = vISheet.getLastRow();
  let dI = vISheet.getRange(1, 1, row, 10).getValues();
  let dRt = vRtSheet.getRange(1, 1, row, 292).getValues();
  let dRn = vRnSheet.getRange(1, 1, row, 148).getValues();
  let dV = vVSheet.getRange(1, 1, row, 148).getValues();
  let dL = vLSheet.getRange(1, 1, row, 148).getValues();
  let dC = vCSheet.getRange(1, 1, row, 148).getValues();

  let src = [...Array(rNum)].map(() => Array(iNum+1));

  for (let i=0; i<rNum; i++) {

    row = tRow;
    for (let j=0; j<iNum; i++) {
      src[i][j] = pSheet.getRange(row, i+3, 1, 1).getValue();
      row += nNum;
    }
    src[i][iNum] = i + 1;

    row = dI.findIndex(x => x[0] === src[i][0]);
    dI = vInfo(row, src[i], dI);
    dRt = vRatio(row, src[i], dRt);
    dRn = vOther(row, src[i], dRn, iNum);
    dV = vOther(row, src[i], dV, 4);
    dL = vOther(row, src[i], dL, 5);
    dC = vOther(row, src[i], dC, 6);
  }

  writeIndicator('v', dI, dRt, dRn, dV, dL, dC);

  row = dI.findIndex(x => x[0] === src[i][0]);
  dI = cISheet.getRange(1, 1, row, 8).getValues();
  dRt = cRtSheet.getRange(1, 1, row, 292).getValues();
  dRn = cRnSheet.getRange(1, 1, row, 148).getValues();
  dV = cVSheet.getRange(1, 1, row, 148).getValues();
  dL = cLSheet.getRange(1, 1, row, 148).getValues();
  dC = cCSheet.getRange(1, 1, row, 148).getValues();
  let dS = cSSheet.getRange(1, 1, row, 99).getValues();
  let dN = cNSheet.getRange(1, 1, row, 99).getValues();
  let dT = cTSheet.getRange(1, 1, row, 99).getValues();

  src.sort(function(a,b){
    if (a[7] > b[7]) { return 1 }
    else { return -1 }
  })

  for (let i=0; i<rNum; i++) {

    let numV = 1, f = true;
    while (f) {
      if (i+1<rNum) {
        if (src[i+1][7] === src[i][7]) {
          i++;
          numV++;
        }
        else { f = false; }
      }
      else { f = false; }
    }

    row = dI.findIndex(x => x[0] === src[i][0]);
    dI = cInfo(row, src.slice(i-numV+1,i+1), dI);
    dRt = cRatio(row, src.slice(i-numV+1,i+1), dRt);
    dRn = cRank(row, src.slice(i-numV+1,i+1), dRn);
    dV = cOther1(row, src.slice(i-numV+1,i+1), dV, 4);
    dL = cOther1(row, src.slice(i-numV+1,i+1), dL, 5);
    dC = cOther1(row, src.slice(i-numV+1,i+1), dC, 6);
    dS = cOther2(row, src.slice(i-numV+1,i+1), dS, 10);
    dN = cOther2(row, src.slice(i-numV+1,i+1), dN, 11);
    dT = cOther2(row, src.slice(i-numV+1,i+1), dT, 12);
  }

  writeIndicator('c', dI, dRt, dRn, dV, dL, dC);

  pSheet.getRange(tRow, 1).setValue('done')
}

function writeIndicator (cat, dI, dRt, dRn, dV, dL, dC) {

  let row = dI.length;
  switch (cat) {
    case 'v':
      vISheet.getRange(1, 1, row, 10).getValues(dI);
      vRtSheet.getRange(1, 1, row, 292).getValues(dRt);
      vRnSheet.getRange(1, 1, row, 148).getValues(dRn);
      vVSheet.getRange(1, 1, row, 148).getValues(dV);
      vLSheet.getRange(1, 1, row, 148).getValues(dL);
      vCSheet.getRange(1, 1, row, 148).getValues(dC);
      return;
    case 'c':
      cISheet.getRange(1, 1, row, 8).getValues(dI);
      cRtSheet.getRange(1, 1, row, 292).getValues(dRt);
      cRnSheet.getRange(1, 1, row, 148).getValues(dRn);
      cVSheet.getRange(1, 1, row, 148).getValues(dV);
      cLSheet.getRange(1, 1, row, 148).getValues(dL);
      cCSheet.getRange(1, 1, row, 148).getValues(dC);
      cSSheet.getRange(1, 1, row, 99).getValues(dS);
      cNSheet.getRange(1, 1, row, 99).getValues(dN);
      cTSheet.getRange(1, 1, row, 99).getValues(dT);
      return;
    default: console.log('■■■■■■■■■■  エラー : writeIndicator  ■■■■■■■■■■')
  }
}

//■■■■■■■■ 動画 ■■■■■■■■

function vInfo(row, src, data) {

  let d = Array(10);
  d[0] = src[0];
  d[1] = src[1];
  d[2] = src[7];
  d[3] = src[8];
  d[4] = src[2];
  d[5] = src[3];
  d[6] = src[13];
  d[7] = src[14];
  d[8] = src[15];
  d[9] = src[16];

  if (row != -1) { data[row] = d }
  else { data.push(d) }

  return data
}

function vRatio(row, src, data) {

}

function vOther(row, src, data, clm) {

  if (row != -1) {
    data[row][0] = src[0];
    data[row][1] = src[1];
    data[row][52+tRow-3] = src[clm];
    //前回との差分
    if (data[row][52+tRow-3-1]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-1]
    }
    //直近1h
    if (data[row][52+tRow-3-2]) {
      data[row][2] = data[row][52+tRow-3] - data[row][52+tRow-3-2]
    }
    //直近24h
    if (data[row][52+tRow-3-48]) {
      data[row][3] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0];
    d[1] = src[1];
    d[52+tRow-3] = src[clm];
    data.push(d);
  }
  return data
}

//  vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle, cDate,
//  sub, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL, cTmb,
//  cCustom, rank

//■■■■■■■■ チャンネル ■■■■■■■■

function cInfo(row, src, data) {

  let d = Array(8);
  d[0] = src[0][7];
  d[1] = src[0][8];
  d[2] = stringify(src.map(x => x[0]));
  d[3] = src[20];
  d[4] = src[9];
  d[5] = src[17];
  d[6] = src[18];
  d[7] = src[19];

  if (row != -1) {
    d[2] = stringify(new Set(d[2].concat(JSON.parse(data[row][2]))));
    data[row] = d;
  }
  else { data.push(d) }

  return data
}

function cRatio(row, src, data) {

}

function cRank(row, src, data) {

  if (row != -1) {
    data[row][0] = src[0][7];
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.sort(function(a,b){return(a[iNum] - b[iNum]);})[0][iNum];
    //前回との差分
    if (data[row][52+tRow-3-1]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-1]
    }
    //直近1h
    if (data[row][52+tRow-3-2]) {
      data[row][2] = data[row][52+tRow-3] - data[row][52+tRow-3-2]
    }
    //直近24h
    if (data[row][52+tRow-3-48]) {
      data[row][3] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.sort(function(a,b){return(a[iNum] - b[iNum]);})[0][iNum];
    data.push(d);
  }
  return data
}

function cOther1(row, src, data, clm) {

  if (row != -1) {
    data[row][0] = src[0][7];
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    //前回との差分
    if (data[row][52+tRow-3-1]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-1]
    }
    //直近1h
    if (data[row][52+tRow-3-2]) {
      data[row][2] = data[row][52+tRow-3] - data[row][52+tRow-3-2]
    }
    //直近24h
    if (data[row][52+tRow-3-48]) {
      data[row][3] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    data.push(d);
  }
  return data
}

function cOther2(row, src, data, clm) {

  if (row != -1) {
    data[row][0] = src[0][7];
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    //直近1h
    if (data[row][52+tRow-3-2]) {
      data[row][2] = data[row][52+tRow-3] - data[row][52+tRow-3-2]
    }
    //直近24h
    if (data[row][52+tRow-3-48]) {
      data[row][3] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    data.push(d);
  }
  return data
}