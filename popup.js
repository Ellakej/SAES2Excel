// Initialize button with user's preferred color
let changeColor = document.getElementById("button");

//chrome.storage.sync.get("color", ({ color }) => {
  //changeColor.style.backgroundColor = color;
//});

// When the button is clicked, inject setPageBackgroundColor into current page
changeColor.addEventListener("click", async () => {
  let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    function: fnExcelReport,
  });
});


// The body of this function will be executed as a content script inside the
// current page
function fnExcelReport()
{
    // Parte 1. Cabeceras
    // Obtencion del RFC y Nombre
    var cab1 = document.getElementById('ctl00_mainCopy_Lbl_Nombre').innerHTML;

    // Obtencion de nombre de materia, grupo y periodo
    var cab2 = document.getElementById('ctl00_mainCopy_GridView1');
    var arrH = cab2.getElementsByTagName("th");
    var arrD = cab2.getElementsByTagName("td");

    // Creacion de la tabla de los datos de cabecera 2
    cab2 = "<tr>" + "<th align=LEFT>" + "Materia" + "</th><td>" + arrD[6].innerHTML + "</td>" + "</tr>" + 
    "<tr>" + "<th align=LEFT>" + "Grupo" + "</th><td>" + arrD[4].innerHTML + "</td>" + 
    "<tr>" + "<th align=LEFT>" + "Semestre" + "</th>" + "<td align=LEFT>" + arrD[0].innerHTML + "</td>" + "</tr>" +"</tr>";

    // Concatenacion al texto principal
    var cab2 = "<table>" + cab2 + "</table>";
    var tab_text = cab1 + cab2;

    // Separador de diseño
    tab_text+="\n<br><table border='2px'><tr bgcolor='#7F1E57' style='color: #FFF;'>";
    var textRange; var j=0;

    // Parte 2. Tabla
    // Obtencion de la tabla de alumnos
    tab = document.getElementsByClassName("ListaAlumnos")[0];
    //rows = tab.getElementsByTagName("tr");                    // Obtenemos todas las filas
    //ans = rows[0].innerHTML;    // Cabeceras de la tabla 2

    //tab_text += ans;

    // Pasar por todos los datos de la tabla
    for(j = 0 ; j < tab.rows.length ; j++) 
    {
        tab_text=tab_text+tab.rows[j].innerHTML+"</tr>";
        
        //tab_text=tab_text+"</tr>";
    }
  

    tab_text=tab_text+"</table>";
    tab_text= tab_text.replace(/<A[^>]*>|<\/A>/g, "");//remove if u want links in your table
    tab_text= tab_text.replace(/<img[^>]*>/gi,""); // remove if u want images in your table
    tab_text= tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE "); 

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // If Internet Explorer
    {
        txtArea1.document.open("txt/html","replace");
        txtArea1.document.write(tab_text);
        txtArea1.document.close();
        txtArea1.focus(); 
        sa=txtArea1.document.execCommand("SaveAs",true,"Excel.xls");
    }  
    else                 //other browser not tested on IE 11
        //sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));  
        var link  = document.createElement('a');
        link.download = "lista_"+arrD[4].innerHTML;
        link.href='data:application/vnd.ms-excel,' + encodeURIComponent(tab_text);
        link.click();

    //return (sa);
}