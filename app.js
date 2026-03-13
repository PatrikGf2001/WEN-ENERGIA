const API_URL="https://script.google.com/macros/s/AKfycbzSBGp3_lQ7EIHmLBzOGWN3hyY-ePiX53oDMQiT_-fIWZ2t-Hj1CPNVmfEcWP4PIe8I/exec"


function mostrarFormulario(tipo){

document.getElementById("impedancia").style.display="none"
document.getElementById("descarga").style.display="none"

document.getElementById(tipo).style.display="block"

}


// AGREGAR FILA
function agregarFila(){

const tabla=document.querySelector("#tabla tbody")

const fila=`
<tr>

<td><input type="number"></td>
<td><input type="number"></td>
<td><input type="number" step="0.01"></td>
<td><input type="number" step="0.01"></td>

<td>

<select>
<option>OK</option>
<option>Crítico</option>
<option>Fin de Vida</option>
</select>

</td>

</tr>
`

tabla.insertAdjacentHTML("beforeend",fila)

}



async function guardarDatos(){

const filas=document.querySelectorAll("#tabla tbody tr")

let datos=[]

filas.forEach(f=>{

const c=f.querySelectorAll("input,select")

datos.push([

document.getElementById("fecha").value,
document.getElementById("cu").value,
document.getElementById("local").value,
document.getElementById("ups").value,
document.getElementById("tipo_bateria").value,
document.getElementById("num_bancos").value,
document.getElementById("celdas_banco").value,
document.getElementById("modelo_equipo").value,
document.getElementById("estado").value,
document.getElementById("temperatura").value,

c[0].value,
c[1].value,
c[2].value,
c[3].value,
c[4].value

])

})

try{

const formData=new URLSearchParams()

formData.append("sheet","IMPEDANCIA")
formData.append("data",JSON.stringify(datos))

await fetch(API_URL,{
method:"POST",
body:formData
})

alert("✅ Su registro fue exitoso")

}catch(error){

alert("Error al guardar")

}

}



function limpiar(){

document.querySelector("#tabla tbody").innerHTML=""
document.querySelectorAll("input").forEach(i=>i.value="")

}



// EXPORTAR PDF
function exportarPDF(){

const elemento=document.getElementById("reporte")

html2pdf().from(elemento).save("reporte_impedancia.pdf")

}
