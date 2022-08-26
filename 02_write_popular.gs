function functionP() {

  if (checkPopular()) { return }

  const srcData1 = getPopular('');
  const srcData2 = getPopular(srcData1[0]);

//  nextPageToken, vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle,
//  cDate, sub, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL,
//  cTmb, cCustom

  let srcData = [...Array(srcData1.length-1)].map(() => []);

  for (let i=0; i<srcData.length; i++) {
    srcData[i] = srcData1[i+1].concat(srcData2[i+1])
  }

  writePopular(srcData);
}

function checkPopular() {
  const time = pSheet.getRange(tRow, 2).getValue();
  const data = pSheet.getRange(tRow, 3).getValue();
  if (data) {
    console.log('実施済み : ' + time);
    return true
  }
  return false
}

function getPopular(nextPageToken) {

  let [
    token, vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle,
    cDate, cntS, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL,
    cTmb, cCustom
  ] = [
    '',[],[],[],[],[],[],[],[],[],
    [],[],[],[],[],[],[],[],[],[],
    [],[]
  ];
  let c_tmp = ''

  const vfields = 'items(id,snippet(title,description,publishedAt,thumbnails(medium(url)),tags,channelId),contentDetails(duration),statistics(viewCount,likeCount,commentCount)),nextPageToken';
  let optJson = {chart: 'mostPopular', regionCode: 'jp', videoCategoryId: vCat, maxResults: 50, fields: vfields, access_token: nextPageToken};

  vJson = YouTube.Videos.list('snippet,contentDetails,statistics',optJson);
  token = vJson.nextPageToken;

  vJson.items.forEach((vJ) => {

    const cfields = 'items(id,snippet(title,description,publishedAt,thumbnails(medium(url)),customUrl),statistics(viewCount,subscriberCount,videoCount)),nextPageToken';
    optJson = {id: vJ.snippet.channelId, fields: cfields};
    cJ = YouTube.Channels.list('snippet,statistics',optJson).items[0];

    vID.push(vJ.id);
    vTitle.push(vJ.snippet.title);
    vDate.push(Utilities.formatDate(new Date(vJ.snippet.publishedAt), 'JST', 'yyyy-MM-dd HH:mm:ss'));
    dur.push(convertTime(vJ.contentDetails.duration));
    cntV.push(vJ.statistics.viewCount);
    cntL.push(vJ.statistics.likeCount);
    cntC.push(vJ.statistics.commentCount);
    cID.push(vJ.snippet.channelId);
    cTitle.push(cJ.snippet.title);
    cDate.push(Utilities.formatDate(new Date(cJ.snippet.publishedAt), 'JST"', 'yyyy-MM-dd HH:mm:ss'));
    cntS.push(cJ.statistics.subscriberCount);
    cntN.push(cJ.statistics.videoCount);
    totV.push(cJ.statistics.viewCount);
    vDesc.push(vJ.snippet.description);
    vURL.push('https://youtube.com/watch?v='+vJ.id);
    vTmb.push(vJ.snippet.thumbnails.medium.url);
    vTags.push(joinArr(vJ.snippet.tags));
    cDesc.push(cJ.snippet.description);
    cTmb.push(cJ.snippet.thumbnails.medium.url);

    c_tmp = removeAt(cJ.snippet.customUrl)
    cCustom.push(c_tmp);

    if (c_tmp) { cURL.push('https://youtube.com/c/'+c_tmp) }
    else { cURL.push('https://youtube.com/channel/'+vJ.snippet.channelId) }
  })

  const srcData = [
    token, vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle,
    cDate, cntS, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL,
    cTmb, cCustom
  ];

  return srcData
}

function writePopular(srcData) {

  let row = tRow;
  for (let i=0; i<srcData.length; i++) {
    pSheet.getRange(row, 3, 1, rNum).setValues([srcData[i]]);
    row += nNum;
  }
}

function joinArr(arr) {

  if (typeof(arr)==='object') { return arr.join(',') }
}

function removeAt(str) {

  if (!str) { return }
  if (str.slice(0,6) === '@user-') { return }
  return str.slice(1)
}

function convertTime(duration) {

  if (!duration || duration==='P0D') { return }
  var reg = new RegExp('^PT([0-9]*H)?([0-9]*M)?([0-9]*S)?');
  var regResult = duration.match(reg);

  var hour = regResult[1];
  var minutes = regResult[2];
  var sec = regResult[3];

  if(hour == undefined) {hour = '00';}
  else {
    hour = hour.split('H')[0];
    if(hour.length == 1){hour = '0' + hour;}
  }

  if(minutes == undefined) {minutes = '00';}
  else {
    minutes = minutes.split('M')[0];
    if(minutes.length == 1){minutes = '0' + minutes;}
  }

  if(sec == undefined) {sec = '00';}
  else {
    sec = sec.split('S')[0];
    if(sec.length == 1){sec = '0' + sec;}
  }

  return hour + ":" + minutes + ":" + sec
}