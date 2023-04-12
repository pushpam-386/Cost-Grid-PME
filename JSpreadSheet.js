const downloadExcel = function() {
    jspreadsheet.render(document.getElementById('spreadsheet'), {
        filename: 'file.xlsx',
    });
}


jspreadsheet.setLicense("eyJhbGciOiJIUzUxMiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjoiMjg3NCIsInVzZXJfc2lnbmF0dXJlIjoiZWE5NzhhYzktNmUzNi00MWYwLTkxNGMtZDBhMzdiMzNiYWU2In0.ZEdadFwXXubQDM64MVhDtC638PuuhZev-K8pozyJMmLQRBBFqbTLh3bs3mUgOgejcZ6KQZu6wbIuFqZe39yC-Q");

var data1 = [
    ['Total', 'Baseline', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['Total', 'Forecast', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['Total', 'Actual', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['Total', 'Variance', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
];


const d = new Date();
var year = d.getFullYear();


function tableData() {
    var table = jspreadsheet(document.getElementById('spreadsheet'),
        {
            worksheets: [{
                data: data1,
                columnDrag: true,
                columns: getCostGridColumns(year),
                mergeCells: {
                    A1: [1, 4]
                }
            }],
        });

    var previous = document.createElement("INPUT");
    previous.setAttribute("type", "button");
    previous.style.border = "white";
    previous.value = '<<';
    previous.addEventListener("click", previousYear);


    var next = document.createElement("INPUT");
    next.setAttribute("type", "button");
    next.style.border = "white";
    next.value = '>>';
    next.addEventListener("click", nextYear);

    function nextYear() {
        year++;
        updateYear(year, previous, next)
        return year;
    }

    function previousYear() {
        year--;
        updateYear(year, previous, next)
        console.log(year);
        return year;
    }
    

    let columnData = document.getElementsByClassName("jss_header")[3];
    columnData.append(previous);
    columnData.append(year);
    columnData.append(next);
    var br = document.createElement("br");
    columnData.append(br);
    columnData.append('FY TOTAL');


    document.getElementById("insert").addEventListener("click", function () {
        let selectionData = document.getElementById("category");
        let val = selectionData.value;
        selectionData.remove(selectionData.selectedIndex);
        var mergeIndex = data1.length + 1;
        

        if (val !== "") {
            var newData = [
                [val, 'Baseline', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                [val, 'Forecast', 0, `=SUM(E${mergeIndex + 1}:P${mergeIndex + 1})`, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                [val, 'Actual', 0, `=SUM(E${mergeIndex + 2}:P${mergeIndex + 2})`, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                [val, 'Variance', 0, `=SUM(E${mergeIndex + 3}:P${mergeIndex + 3})`, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            ]

            console.log(table)
            console.log("merge", `A${mergeIndex}`);
            data1 = data1.concat(newData);
            table[0].setMerge(`A${mergeIndex}`, 1, 4);
            table[0].loadData(data1);
        }
        console.log(data1);
    })
}

tableData();

function updateYear(year, previous, next) {
    let columnData = document.getElementsByClassName("jss_header")[3];
    columnData.innerText = '';
    columnData.append(previous);

    columnData.append(year);
    columnData.append(next);
    var br = document.createElement("br");
    columnData.append(br);
    columnData.append('FY TOTAL');
}


function getCostGridColumns(year) {
    var headerColumns = [
        {
            type: 'text',
            width: '200',
            title: 'CATEGORY',
            // align: 'left',
            // wordWrap: true,
            readOnly: true
        },
        {
            type: 'text',
            width: '100',
            title: 'TYPE',
            // align: 'left',
            // wordWrap: true,
            readOnly: true
        },
        {
            type: 'text',
            width: 150,
            title: 'LIFE TIME TOTAL',
            // align: 'right',
            // wordWrap: true,
            readOnly: true,
            // //	mask:'#,##,00', decimal:',',
            // font: 'Verdana,Arial,sans-serif'

        },
        {
            type: 'text',
            width: '120',
            title: ' ',
            // align: 'center',
            // wordWrap: true,
            readOnly: true,
            // //mask:'#,##,00', decimal:',' 						

        },
    ];


    var months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    for (var i = 0; i < 12; i++) {

        headerColumns.push(
            {
                type: 'number',
                width: '90',
                id: 'month',
                title: months[i] + ' ' + (year),
                align: 'right',
                wordWrap: true,
                // readOnly:false,
                // editable: true,

            }
        );
    }
    return headerColumns;
}


document.getElementById('addCategory').addEventListener('click', openForm1);

document.getElementById('crossImage1').addEventListener('click', closeForm1);

document.getElementById('manageCategory').addEventListener('click', openForm2);

document.getElementById('crossImage2').addEventListener('click', closeForm2);

document.getElementById('downloadExcelData').addEventListener('click', downloadExcel);



function openForm1() {
    document.getElementById("addCategoryPopup").style.display = "block";
}

function closeForm1() {
    document.getElementById("addCategoryPopup").style.display = "none";
}


function openForm2() {
    document.getElementById("manageCategoryPopup").style.display = "block";
}

function closeForm2() {
    document.getElementById("manageCategoryPopup").style.display = "none";
}



// Add-on for Jspreasheet
jspreadsheet.setExtensions({ render });




