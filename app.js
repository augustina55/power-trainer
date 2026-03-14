import { Chess } from "https://cdn.jsdelivr.net/npm/chess.js@1.0.0/dist/esm/chess.js";
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

let tableData=[]
let excelRows=[]
let idCounter=351

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

const chapterSelect=document.getElementById("chapterSelect")

chapters.forEach(c=>{
let option=document.createElement("option")
option.value=c.id
option.textContent=c.title
chapterSelect.appendChild(option)
})

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

document.getElementById("generateBtn").addEventListener("click",generateRows)
document.getElementById("downloadBtn").addEventListener("click",downloadExcel)

function generateRows(){

let puzzleType=document.getElementById("puzzleType").value
let chapterId=parseInt(document.getElementById("chapterSelect").value)

excelRows.forEach(row=>{

let title=row.title
let pgn=row.pgn

let chess=new Chess()

chess.loadPgn(pgn)

let headers=chess.header()

let fen=headers.FEN

let movesSAN=chess.history()

chess.reset()
chess.load(fen)

let moves=[]

movesSAN.forEach((m,i)=>{

let move=chess.move(m)

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

let trainerRow={

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

tableData.push(trainerRow)
addRow(trainerRow)

})

}

function addRow(row){

let tbody=document.querySelector("#resultTable tbody")

let tr=document.createElement("tr")

Object.values(row).forEach(v=>{

let td=document.createElement("td")
td.textContent=v
tr.appendChild(td)

})

tbody.appendChild(tr)

}

function downloadExcel(){

let ws=XLSX.utils.json_to_sheet(tableData)

let wb=XLSX.utils.book_new()

XLSX.utils.book_append_sheet(wb,ws,"trainer")

XLSX.writeFile(wb,"power_trainer_output.xlsx")

}
