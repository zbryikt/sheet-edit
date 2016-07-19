
is-loading = ->
  document.querySelector \.loader .className = (if it => <[loading]> else []).concat(<[loader]>).join(" ")
data = do
  rows: []
  headers: []
  trs: []
  clusterizer: null

tocsv = ->
  headers = data.headers
  ret = data.rows.map((d) -> headers.map(-> "#{d[it]}").join(\,)).join(\\n)
  ret = "#{headers.map(->'"' + it + '"').join(\,)}\n#{ret}"
  url = URL.createObjectURL new Blob [ret], {type: \text/csv}
  link = document.getElementById \download-link
  link.setAttribute \href, url
  link.style.display = \block

fields = ->
  headers = data.headers
  fields = headers.map -> ret = name: it, data: []
  for i from 0 til data.rows.length =>
    for j from 0 til headers.length =>
      fields[j].data.push data.rows[i][headers[j]]

render = ->
  head = document.querySelector '#sheet .sheet-head'
  scroll = document.querySelector '#sheet .clusterize-scroll'
  content = document.querySelector '#sheet .clusterize-content'
  content.innerHTML = ""
  h = data.headers
  w = "#{100/h.length}%"
  if  h.length > 10 => w = "10%"
  trs = data.rows.map (row,i) -> (
    "<div>" + h.map((d,j)-> 
      "<div contenteditable='true' row='#i' col='#j' style='width:#w'>#{row[d] or ''}</div>"
    ).join("") + "</div>"
  )
  head.innerHTML = "<div>" + data.headers.map(-> "<div style='width:#w'>#it</div>").join("") + "</div>" 
  if data.clusterizer => that.destroy true
  data.clusterizer = new Clusterize do
    rows: trs
    scrollElem: scroll
    contentElem: content
  is-loading false
  scroll.addEventListener \scroll, (e) -> head.scrollLeft = scroll.scrollLeft

  content.addEventListener \click, (e) ->
    setTimeout (->
      n = e.target
      row = +n.getAttribute(\row)
      col = +n.getAttribute(\col)
      d = data.rows[row][data.headers[col]]
      document.getElementById(\input).value = d
    ), 0

  content.addEventListener \keydown, (e) ->
    setTimeout (->
      n = e.target
      val = n.textContent
      row = +n.getAttribute(\row)
      col = +n.getAttribute(\col)
      data.rows[row][data.headers[col]] = d = val
      document.getElementById(\input).value = d
    ), 0

sheet = (ret) ->
  list = ret.result.values
  data.headers = h = list.0
  list.splice 0,1
  data.rows = list.map (v) -> 
    hash = {}
    for i from 0 til v.length => hash[h[i]] = v[i]
    hash
  console.log "[Sheet] #{data.rows.length} rows imported."
  render!

csv = (buf) ->
  is-loading true
  Papa.parse buf, do
    worker: true, header: true
    step: ({data: rows}) ->
      data.rows.push.apply data.rows, rows
    complete: ->
      data.headers = [k for k of data.rows.0 or {}]
      console.log "[CSV] #{data.rows.length} rows imported."
      render!

xls = (buf) ->
  is-loading true
  workbook = XLSX.read buf, {type: \binary}
  sheet = workbook.Sheets[workbook.SheetNames[0]]
  list = XLSX.utils.sheet_to_json(sheet)
  data.headers = h = [k for k of (list[0] or {})]
  data.rows = list
  data.rows.map (row) -> for k of h => if !(row[k]?) => row[k] = ""
  console.log "[Excel] #{list.length} rows imported."
  render!

loadfile = ->
  file = document.getElementById \file
  if !file.files.length => return
  fr = new FileReader!
  fr.onload = ->
    data <<< rows: [], headers: [], trs: []
    [name,buf] = [file.files.0.name, fr.result]
    if /\.csv$/.exec name => csv buf
    if /\.xlsx?$/.exec name => xls buf
  fr.readAsBinaryString file.files.0
