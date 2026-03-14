let excelRows=[]
let tableData=[]
let idCounter=1

const chapters=[

{ id:270,title:"Test"},
{ id:317,title:"The chessboard"},
{ id:324,title:"Meet the Bishop"},
{ id:325,title:"Meet the queen"},
{ id:326,title:"Meet the king"},
{ id:331,title:"Meet the pawn"},
{ id:336,title:"Understanding Check"},
{ id:337,title:"How to Get Out of Check"},
{ id:323,title:"Meet the Rook"},
{ id:338,title:"Checkmate"},
{ id:339,title:"Stalemate"},
{ id:353,title:"Ladder Rook Checkmate"},
{ id:333,title:"Understanding the attack"},
{ id:334,title:"Understanding defense"},
{ id:335,title:"Captures in chess"},
{ id:341,title:"Checkmate in one move"},
{ id:344,title:"Castling"},
{ id:351,title:"Assisted checkmate basics"},
{ id:342,title:"checkmate with a queen"},
{ id:343,title:"Ladder Checkmate"}

]

/* populate chapter dropdown */

let chapterSelect=document.getElementById("chapterSelect")

chapters.forEach(c=>{

let option=document.createElement("option")
option.value=c.id
option.textContent=c.title

chapterSelect.appendChild(option)

})

/* read excel */

document.getElementById("excelFile").addEventListener("change",function(e){

let file=e.target.files[0]

let reader=new FileReader()

reader.onload=function(evt){

let data=new Uint8Array(evt.target.result)

let workbook=XLSX.read(data,{type:"array"})

let sheet=workbook.Sheets[workbook.SheetNames[0]]

excelRows=XLSX.utils.sheet_to_json(sheet)

}

reader.readAsArrayBuffer(file)

})

/* generate trainer rows */

function generateRows(){

tableData=[]

let puzzleType=document.getElementById("puzzleType").value
let chapterId=parseInt(document.getElementById("chapterSelect").value)

excelRows.forEach(row=>{

let title=row.title||row.Title
let pgn=row.pgn||row["PGN Text"]

if(!title||!pgn)return

pgn=pgn.trim()

/* extract FEN */

let fenMatch=pgn.match(/\[FEN\s+"([^"]+)"\]/)

if(!fenMatch)return

let fen=fenMatch[1]

let fenParts=fen.split(" ")
fenParts[4]="0"
fenParts[5]="1"
fen=fenParts.join(" ")

/* extract moves */

let movesSection=pgn.split("\n\n").pop()

movesSection=movesSection.replace(/\{[^}]*\}/g,"")
movesSection=movesSection.replace(/\[%[^\]]*\]/g,"")
movesSection=movesSection.replace(/1-0|0-1|1\/2-1\/2|\*/g,"")
movesSection=movesSection.replace(/\d+\.+/g,"")

let sanMoves=movesSection.trim().split(/\s+/)

let chess=new Chess()

chess.load(fen)

let moves=[]

sanMoves.forEach((m,i)=>{

let move=chess.move(m)

if(!move)return

moves.push({

ply:i+1,
hint:null,
move:move.from+move.to,
arrows:[],
comment:"",
variations:[],
annotations:[],
highlighted_squares:[]

})

})

let moveJSON=JSON.stringify(moves)

let now=new Date().toISOString()

let rowData={

id:idCounter++,
title:title,
fen:fen,
puzzle_type:puzzleType,
config:"{}",
custom_pieces:null,
moves:moveJSON,
arrows:null,
highlighted_squares:null,
board_disable:false,
points:0,
position_order:1,
created_at:now,
updated_at:now,
chapter_id:chapterId,
unlock_after_page_id:null,
top_level_hint:null

}

tableData.push(rowData)

addRow(rowData)

})

}

/* add table row */

function addRow(data){

let table=document.querySelector("#resultTable tbody")

let tr=document.createElement("tr")

tr.innerHTML=`

<td>${data.id}</td>
<td>${data.title}</td>
<td>${data.fen}</td>
<td>${data.puzzle_type}</td>
<td>${data.moves}</td>
<td>${data.chapter_id}</td>

`

table.appendChild(tr)

}

/* download excel */

function downloadExcel(){

let ws=XLSX.utils.json_to_sheet(tableData)

let wb=XLSX.utils.book_new()

XLSX.utils.book_append_sheet(wb,ws,"trainer")

XLSX.writeFile(wb,"power_trainer.xlsx")

}
