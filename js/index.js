
$(document).ready(function () {
    var gData = [];
    var gNewData = [];

    $('button#chooseFile.btn').click(function (e) { 
        $('input#fileUpload').click();
    });

    $('input[type="file"]').change(function(e){
        var fileName = e.target.files[0].name;
    });


    $("body").on("click", "#upload", function () {
        //Reference the FileUpload element.
        var fileUpload = $("#fileUpload")[0];

        //Validate whether File is valid Excel file.
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
        if (regex.test(fileUpload.value.toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();

                //For Browsers other than IE.
                if (reader.readAsBinaryString) {
                    reader.onload = function (e) {
                        ProcessExcel(e.target.result);
                    };
                    reader.readAsBinaryString(fileUpload.files[0]);
                } else {
                    //For IE Browser.
                    reader.onload = function (e) {
                        var data = "";
                        var bytes = new Uint8Array(e.target.result);
                        for (var i = 0; i < bytes.byteLength; i++) {
                            data += String.fromCharCode(bytes[i]);
                        }

                        ProcessExcel(data);
                    };
                    reader.readAsArrayBuffer(fileUpload.files[0]);
                }
            } else {
                alert("This browser does not support HTML5.");
            }
        } else {
            alert("Please upload a valid Excel file.");
        }
    });

    function ProcessExcel(data) {
        
        //Read the Excel File data.
        var workbook = XLSX.read(data, {
            type: 'binary'
        });

        //Fetch the name of First Sheet.
        var firstSheet = workbook.SheetNames[0];

        //Read all rows from First Sheet into an JSON array.
        var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);

        for(var obj of excelRows) {
            gData.push(obj);  
        }

        //Create a HTML Table element.
        var table = $("<table class='table table-hover text-center' />");
        table[0].border = "1";

        //Add the header row.
        var row = $(table[0].insertRow(-1));
        
        //Add the header cells.
        var headerCell = $("<th class='text-' />");
        headerCell.html("Fecha de comprobante");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Tipo de Cfe");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Serie");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Número");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Tipo de moneda");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Tipo de cambio");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Monto total a pagar");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Documento del receptor");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Razón social del receptor");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Dirección del receptor");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Ciudad del receptor");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Departamento del receptor");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("Adenda");
        row.append(headerCell);

        //Add the data rows from Excel file.
        for (var i = 0; i < excelRows.length; i++) {
            // console.log(excelRows[i]);
            //Add the data row.
            var row = $(table[0].insertRow(-1));

            //Add the data cells.
            var cell = $("<td />");
            cell.html(excelRows[i]["Fecha de comprobante"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Tipo de Cfe"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Serie"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(excelRows[i]["Número"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Tipo de moneda"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Tipo de cambio"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(excelRows[i]["Monto total a pagar"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Documento del receptor"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Razón social del receptor"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(excelRows[i]["Dirección del receptor"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Ciudad del receptor"]);
            row.append(cell);
    
            var cell = $("<td />");
            cell.html(excelRows[i]["Departamento del receptor"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(excelRows[i]["Adenda"]);
            row.append(cell);
        }

        var dvExcel = $("#dvExcel");
        dvExcel.html("");
        dvExcel.append(table);
    };

    $('button#newFilePreview').click(function(e) {
        $('button#generateFile').removeAttr('hidden');
        $('button#newFilePreview').attr('hidden', 'true');

        // console.log(gData);
        Object.size = function(obj) {
            var size = 0, key;
            for (key in obj) {
                if (obj.hasOwnProperty(key)) size++;
            }
            return size;
        };

        for(var obj of gData) {
            
        }

        for (let i = 0; i < gData.length; i++) {
            if(Object.size(gData[i]) > 5) {
                var newRow = {
                    "BOLETA N°": gData[i]["Serie"] + gData[i]["Número"],
                    "DETALLE DE OBJETO": gData[i]["Adenda"],
                    "FECHA": gData[i]["Fecha de comprobante"],
                    "PRECIO PAGADO": gData[i]["Tipo de moneda"] + gData[i]["Monto total a pagar"],
                    "NOMBRES Y APELLIDOS": gData[i]["Razón social del receptor"],
                    "DOMICILIO / N°": gData[i]["Dirección del receptor"] + "," + gData[i]["Ciudad del receptor"] + "," + gData[i]["Departamento del receptor"],
                    "DOCUMENTO": gData[i]["Documento del receptor"]
                };
                gNewData.push(newRow);
            }
        }


        //Create a HTML Table element.
        var table = $("<table class='table table-hover text-center' />");
        
        table[0].border = "1";
        table.append(`<tr>
                        <td class="font-weight-bold">
                            <h1>
                            Departamento de Contralor Social 
                            </h1>
                        </td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>

                    <tr>
                        <td></td>
                        <td class="font-weight-bold">
                            <h4>
                            Jefatura  V  Especializada
                            </h4>
                        </td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>

                    <tr>
                        <td></td>
                        <td>
                            <h4>
                            Carlos Quijano 1310 Tel- 152 2731  CP 11100
                            </h4>
                        </td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td class="font-weight-bold">
                            <h3>
                            E-mail: dit-compraventa.montevideo@minterior.gub.uy
                            </h3>
                        </td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td></td>
                        <td class="font-weight-bold">
                            <h1>
                            DECLARACION DE COMPRA VENTA
                            </h1>
                        </td>  
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>

                    <tr>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td class="font-weight-bold">
                            <h1>
                            1158
                            </h1>
                        </td>
                        <td></td>
                    </tr>

                    <tr>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>COMERCIO:</td> 
                        <td>RUBRO: Joyería y Anticuario</td>
                        <td></td>
                        <td>DIRECCION: Sarandí 635</td>
                        <td></td>
                        <td></td>
                    </tr>`);
        //Add the header row.
        var row = $(table[0].insertRow(-1));
        
        //Add the header cells.
        var headerCell = $("<th />");
        headerCell.html("BOLETA N°");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("DETALLE DE OBJETO");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("FECHA");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("PRECIO PAGADO");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("NOMBRES Y APELLIDOS");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("DOMICILIO / N°");
        row.append(headerCell);

        var headerCell = $("<th />");
        headerCell.html("DOCUMENTO");
        row.append(headerCell);
        
        //Add the data rows from Excel file.
        for (var i = 0; i < gNewData.length; i++) {
            //Add the data row.
            var row = $(table[0].insertRow(-1));
                    
            //Add the data cells.
            var cell = $("<td />");
            cell.html(gNewData[i]["BOLETA N°"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(gNewData[i]["DETALLE DE OBJETO"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(gNewData[i]["FECHA"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(gNewData[i]["PRECIO PAGADO"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(gNewData[i]["NOMBRES Y APELLIDOS"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(gNewData[i]["DOMICILIO / N°"]);
            row.append(cell);

            var cell = $("<td />");
            cell.html(gNewData[i]["DOCUMENTO"]);
            row.append(cell);
        }
        
        var dvExcel = $("#dvExcel");
        dvExcel.html("");
        dvExcel.append(table);
    });

    $('button#generateFile').click(function(e) {
        $("div#dvExcel").table2excel({
            // exclude CSS class
            exclude: ".noExl",
            name: "Worksheet Name",
            filename: "output", //do not include extension
            fileext: ".xls" // file extension
        });
    });
});