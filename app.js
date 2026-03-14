import { Chess } from "https://cdn.jsdelivr.net/npm/chess.js@1.0.0/dist/esm/chess.js";
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

let tableData=[]
let idCounter=351


document.getElementById("fileInput").addEventListener("change",e=>{

let file=e.target.files[0]

let reader=new FileReader()

reader.onload=ev=>{
document.getElementById("pgnInput").value=ev.target.result
}

reader.readAsText(file)

})


document.getElementById("generateBtn").addEventListener("click",createTrainer)
document.getElementById("downloadBtn").addEventListener("click",downloadExcel)



function createTrainer(){

let pgn=document.getElementById("pgnInput").value

let chess=new Chess()

chess.loadPgn(pgn)

let headers=chess.header()

let fen=headers.FEN

if(!fen){

alert("FEN not found in PGN")
return

}

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


let row={

id:idCounter++,
title:"Stalemate",
fen:fen,
puzzle_type:"single_variant",
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
chapter_id:339,
unlock_after_page_id:null,
top_level_hint:null

}


tableData.push(row)

addRow(row)

}



function addRow(row){

let tbody=document.querySelector("#resultTable tbody")

let tr=document.createElement("tr")

Object.values(row).forEach(val=>{

let td=document.createElement("td")
td.textContent=val
tr.appendChild(td)

})

tbody.appendChild(tr)

}



function downloadExcel(){

let ws=XLSX.utils.json_to_sheet(tableData)

let wb=XLSX.utils.book_new()

XLSX.utils.book_append_sheet(wb,ws,"puzzles")

XLSX.writeFile(wb,"power_trainer.xlsx")

}