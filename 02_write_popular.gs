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
/*
function functionP() {

  let nextPageToken = '';
  let jsonData = getPopular(nextPageToken);
  const [vID_1, vTitle_1, pDate_1, dur_1, cntV_1, cntL_1, cntC_1, cID_1] = setPopular(jsonData);

  nextPageToken = jsonData.nextPageToken;
  jsonData = getPopular(nextPageToken);
  const [vID_2, vTitle_2, pDate_2, dur_2, cntV_2, cntL_2, cntC_2, cID_2] = setPopular(jsonData);

  writePopular([vID_1.concat(vID_2)]);
  writePopular([vTitle_1.concat(vTitle_2)]);
  writePopular([vDate_1.concat(vDate_2)]);
  writePopular([dur_1.concat(dur_2)]);
  writePopular([cntV_1.concat(cntV_2)]);
  writePopular([cntL_1.concat(cntL_2)]);
  writePopular([cntC_1.concat(cntC_2)]);
  writePopular([cID_1.concat(cID_2)]);
}
*/

function checkPopular() {
  const time = pSheet.getRange(pRow, 2).getValue();
  const data = pSheet.getRange(pRow, 3).getValue();
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

  for (let i=0; i<srcData.length; i++) {
    pSheet.getRange(pRow, 3, 1, 100).setValues([srcData[i]]);
    pRow += 100;
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