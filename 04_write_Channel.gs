function functionC() {
  if (pSheet.getRange(tRow, 1).getValue() != 'video') { return }

  let row = cISheet.getLastRow();
  let cI = cISheet.getRange(1, 1, row, 8).getValues();
  let cRt = cRtSheet.getRange(1, 1, row, 196).getValues();
  let cV = cVSheet.getRange(1, 1, row, 148).getValues();
  let cL = cLSheet.getRange(1, 1, row, 148).getValues();
  let cC = cCSheet.getRange(1, 1, row, 148).getValues();
  let cS = cSSheet.getRange(1, 1, row, 99).getValues();
  let cN = cNSheet.getRange(1, 1, row, 99).getValues();
  let cT = cTSheet.getRange(1, 1, row, 99).getValues();

  let cRn = cRnSheet.getRange(2, 1, row-1, 148).getValues();
  let arr = cRnSheet.getRange(1, 1, 1, 148).getDisplayValues()[0];
  cRn.unshift(arr);

  let cP = cPSheet.getRange(2, 1, row, 196).getValues();
  arr = cRnSheet.getRange(1, 1, 1, 196).getDisplayValues()[0];
  cP.unshift(arr);

  for (let i=1; i<cI.length; i++) {
    cRt[i][2] = '', cRt[i][3] = '', cRt[i][52+tRow-3] = cRt[i][52+tRow-3-1];
    cV[i][2] = '', cV[i][3] = '';
    cL[i][2] = '', cL[i][3] = '';
    cC[i][2] = '', cC[i][3] = '';
    cS[i][2] = '', cS[i][3] = '';
    cN[i][2] = '', cN[i][3] = '';
    cT[i][2] = '', cT[i][3] = '';
  }

  row = cISheet.getLastRow();
  const vRt = vRtSheet.getRange(1, 1, row, 196).getDisplayValues();

  let src = Array(iNum+1);

  row = tRow;
  for (let i=0; i<iNum; i++) {
    src[i] = pSheet.getRange(row, 3, 1, rNum).getValues()[0];
    row += nNum;
  }
  src[iNum] = pSheet.getRange(1, 3, 1, rNum).getValues()[0];
  src = src[0].map((_, c) => src.map(r => r[c]));

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

    row = cI.findIndex(x => x[0] === src[i][7]);
    cI = cInfo(row, src.slice(i-numV+1,i+1), cI);
    cRn = cRank(row, src.slice(i-numV+1,i+1), cRn);
    cRt = cRatio(row, src.slice(i-numV+1,i+1), vRt, cRt);
    cP = cPeriod(row, src.slice(i-numV+1,i+1), cP);
    cV = cOther1(row, src.slice(i-numV+1,i+1), cV, 4);
    cL = cOther1(row, src.slice(i-numV+1,i+1), cL, 5);
    cC = cOther1(row, src.slice(i-numV+1,i+1), cC, 6);
    cS = cOther2(row, src.slice(i-numV+1,i+1), cS, 10);
    cN = cOther2(row, src.slice(i-numV+1,i+1), cN, 11);
    cT = cOther2(row, src.slice(i-numV+1,i+1), cT, 12);
  }

  writeIndicator('c', cI, cRn, cRt, cP, cV, cL, cC, cS,cN, cT);

  pSheet.getRange(tRow, 1).setValue('channel')
}

//  vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle, cDate,
//  sub, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL, cTmb,
//  cCustom, rank

//■■■■■■■■ チャンネル ■■■■■■■■

function cInfo(row, src, data) {

  let d = Array(8);
  d[0] = src[0][7];
  d[1] = src[0][8];
  d[2] = JSON.stringify(src.map(x => x[0]));
  d[3] = src[0][20];
  d[4] = src[0][9];
  d[5] = src[0][17];
  d[6] = src[0][18];
  d[7] = src[0][19];

  if (row != -1) {
    d[2] = JSON.stringify(new Set(d[2].concat(JSON.parse(data[row][2]))));
    data[row] = d;
  }
  else { data.push(d) }

  return data
}

function cRank(row, src, data) {

  if (row != -1) {
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.sort(function(a,b){return(a[iNum] - b[iNum]);})[0][iNum];
    //前日との差分
    if (data[row][52+tRow-3-48]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
    //直近1h
    if (data[row][52+tRow-3] < data[row][2]) {
      data[row][2] = data[row][52+tRow-3];
      data[row][3] = rDay + data[0][52+tRow-3];
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.sort(function(a,b){return(a[iNum] - b[iNum]);})[0][iNum];
    d[2] = d[52+tRow-3];
    d[3] = rDay + data[0][52+tRow-3];
    data.push(d);
  }
  return data
}

function cRatio(row, src, vRt, data) {

  sumR = vRt.filter(function(v){return v[0] === src[0][0];}).reduce((sum, x) => sum + x[148+tRow-3], 0);

  if (row != -1) {
    data[row][1] = src[0][8];
    data[row][148+tRow-3] = sumR;
    data[row][52+tRow-3] = data[row][52+tRow-3-1] + sumR;
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
    d[148+tRow-3] = sumR;
    d[52+tRow-3] = sumR;
    data.push(d);
  }
  return data
}

function cPeriod(row, src, data) {

  if (row != -1) {
    data[row][1] = src[0][8];
    //ランクイン開始日時
    if (data[row][148+tRow-3-1]) { data[row][148+tRow-3] = data[row][148+tRow-3-1] }
    else { data[row][148+tRow-3] = rDay + data[0][148+tRow-3] }
    //ランクイン期間
    if (data[row][52+tRow-3-1]) { data[row][52+tRow-3] = data[row][52+tRow-3-1] + 0.5 }
    else { data[row][52+tRow-3] = 0.5 }
    //最長期間の更新
    if (data[row][52+tRow-3] > data[row][2]) {
      data[row][2] = data[row][52+tRow-3];
      data[row][3] = data[row][52+tRow-3] + '～' + rDay + data[0][52+tRow-3];
    }
  }
  else {
    let d = Array(196);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[2] = 0.5;
    d[3] = rDay + data[0][52+tRow-3] + '～' + rDay + data[0][52+tRow-3];
    d[52+tRow-3] = 0.5;
    d[148+tRow-3] = rDay + data[0][148+tRow-3];
    data.push(d);
  }
  return data
}

function cOther1(row, src, data, clm) {

  if (row != -1) {
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
    let d = Array(99);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    data.push(d);
  }
  return data
}