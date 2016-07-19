api-key = \AIzaSyD3emlU63t6e_0n9Zj9lFCl-Rwod0OMTqY
client-id = \1003996266757-4gv30no8ije0sd8d8qsd709dluav0676.apps.googleusercontent.com

scopes=  <[
  profile
  https://www.googleapis.com/auth/drive.metadata.readonly
  https://www.googleapis.com/auth/spreadsheets.readonly
]>.join(' ')

init = ->
  gapi.client.set-api-key api-key
  gapi.auth2.init do
    client_id: client-id
    scope: scopes

pageToken = null
sheet-files = []

listFiles = (token) ->
  config = do
    pageSize: 20,
    fields: "nextPageToken, files(id, name)"
    q: "mimeType='application/vnd.google-apps.spreadsheet'"
  if token => config.pageToken = token
  request = gapi.client.drive.files.list config
  request.execute (ret) ->
    pageToken := ret.nextPageToken
    if ret.files =>
      list = document.getElementById \sheet-list
      ret.files.map ->
        node = document.createElement \div
        node.innerHTML = it.name
        node.sheetId = it.id
        list.appendChild(node)
      sheet-files.push.apply sheet-files, ret.files

loadFile = (id) ->
  is-loading true
  gapi.client.load 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    .then ->
      gapi.client.sheets.spreadsheets.values.get do
        spreadsheetId: id
        range: 'A:ZZ'
    .then -> sheet it


handle = ->
  gapi.auth2.get-auth-instance!sign-in!then ->
    gapi.client.load 'drive', 'v3', listFiles
    list = document.getElementById \sheet-list
    list.style.display = \block
    list.addEventListener \click, (e) ->
      list.style.display = \none
      if e.target.sheetId => loadFile that

gapi.load 'client:auth2', init
