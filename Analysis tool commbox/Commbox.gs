const COMMBOX_TOKEN = PROPS.getProperty("COMMBOX_TOKEN");

function fetchTranscriptFromCommBox(streamId, conversationId) {
  const url = `https://api.commbox.io/streams/${streamId}/objects/${conversationId}/transcript`;

  const options = {
    method: 'get',
    headers: { 'Authorization': `Bearer ${COMMBOX_TOKEN}` },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const data = json.data;

    if (!data) {
      logAction("CommBox Error", "No data returned in transcript response", conversationId, streamId);
      return null;
    }
    if (typeof data === 'object' && !Array.isArray(data) && data.transcript) {
      logAction("CommBox Success", "Transcript retrieved (object)", conversationId, streamId);
      return data.transcript;
    }
    if (Array.isArray(data)) {
      for (let item of data) {
        if (item.transcript) {
          logAction("CommBox Success", "Transcript retrieved (array)", conversationId, streamId);
          return item.transcript;
        }
      }
    }
    logAction("CommBox Error", "Transcript not found in data", conversationId, streamId);

  } catch (err) {
    logAction("CommBox Exception", "Failed to fetch transcript: " + err, conversationId, streamId);
  }

  return null;
}