const excel_file = document.getElementById("excel_file");

excel_file.addEventListener("change", (event) => {
  var count = 0;

  if (
    ![
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
    ].includes(event.target.files[0].type)
  ) {
    document.getElementById("excel_data").innerHTML =
      '<div class="alert alert-danger">Only .xlsx or .xls file format are allowed</div>';

    excel_file.value = "";

    return false;
  }

  var reader = new FileReader();

  reader.readAsArrayBuffer(event.target.files[0]);

  reader.onload = function (event) {
    var data = new Uint8Array(reader.result);

    var work_book = XLSX.read(data, { type: "array" });

    var sheet_name = work_book.SheetNames;

    var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {
      header: 1,
    });

    if (sheet_data.length > 0) {
      var table_output = '<table class="table table-striped table-hover" id="table-data">';

      for (var row = 0; row < sheet_data.length; row++) {
        table_output += "<tr>";

        for (var cell = 0; cell <= sheet_data[row].length; cell++) {
          if (row == 0) {
            if(cell != sheet_data[row].length) 
            {
                table_output += "<th>" + sheet_data[row][cell] + "</th>";
            }
            else 
            {
                table_output += "<th>" + "" + "</th>";
            }
          } 
          else if(cell == sheet_data[row].length) {
            table_output += "<td>" + '<a href="#deleteEmployeeModal" onclick="delRow(this);" class="delete" data-toggle="modal"><i class="material-icons" data-toggle="tooltip" title="Sil">&#xE872;</i></a>' + "</td>";
          }
          else {
            table_output += "<td>" + sheet_data[row][cell] + "</td>";
          }
        }

        table_output += "</tr>";
        count += 1;
      }

      table_output += "</table>";

      document.getElementById("excel_data").innerHTML = table_output;
      document.getElementById("hidden-counter").style = "display: block";
      document.getElementById("counter").innerHTML = count - 1;
    }

    excel_file.value = "";
  };
});

function delRow(btn){
    if(!confirm("Kayıt silinsin mi?")) return;
    
    var tbl = btn.parentNode.parentNode.parentNode;
    var row = btn.parentNode.parentNode.rowIndex;

    tbl.deleteRow(row);

    var table = document.getElementById("table-data");
    var tbodyRowCount = table.tBodies[0].rows.length;
    document.getElementById("counter").innerHTML = tbodyRowCount - 1;
}