<html>

<head>
    <title>Cost Grid</title>

    <!-- <script src="https://jspreadsheet.com/v9/jspreadsheet.js"></script>
    <script src="https://jsuites.net/v4/jsuites.js"></script>
    <link rel="stylesheet" href="https://jspreadsheet.com/v9/jspreadsheet.css" type="text/css" />
    <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" />
    <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Material+Icons" />

    <script src="https://cdn.jsdelivr.net/npm/jszip@3.6.0/dist/jszip.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@jspreadsheet/render/dist/index.min.js"></script> -->

    <script src="https://bossanova.uk/jspreadsheet/v4/jexcel.js"></script>
    <script src="https://jsuites.net/v4/jsuites.js"></script>
    <link rel="stylesheet" href="https://bossanova.uk/jspreadsheet/v4/jexcel.css" type="text/css" />
    <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" />

    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jspreadsheet-ce/dist/jspreadsheet.min.css"
        type="text/css" />
    <script type="text/javascript" src="https://cdn.jsdelivr.net/npm/jspreadsheet-ce/dist/index.min.js"></script>

    <!-- <script src="https://jspreadsheet.com/v9/jspreadsheet.js"></script> -->
    <!-- <script src="https://jsuites.net/v4/jsuites.js"></script> -->
    <!-- <link rel="stylesheet" href="https://jspreadsheet.com/v9/jspreadsheet.css" type="text/css" /> -->
    <!-- <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" /> -->


    <!-- <script src="https://bossanova.uk/jspreadsheet/v4/jexcel.js"></script>
    <link rel="stylesheet" href="https://bossanova.uk/jspreadsheet/v4/jexcel.css" type="text/css" />
    <script src="https://jsuites.net/v4/jsuites.js"></script>
    <link rel="stylesheet" href="https://jsuites.net/v4/jsuites.css" type="text/css" />  -->

    <!-- <link rel="stylesheet" type="text/css" href="styles.css" /> -->




    <style>
        #category {
            border: 1px solid black;
            width: 420px;
            height: 40px;
            text-align: center;
        }

        #insert {
            border: 2px solid black;
            width: 90px;
            height: 40px;
            color: aliceblue;
            background-color: rgb(100, 100, 183);
        }

        #spreadsheet {
            margin-top: 20px;
        }

        #addcategory {
            margin-left: 50px;
            margin-top: 80px;
            border: 1px solid black;
            width: 120px;
            height: 40px;
            color: rgb(0, 0, 0);
            background-color: rgb(236, 236, 241);
        }

        #manageCategory {
            margin-top: 80px;
            border: 1px solid black;
            width: 150px;
            height: 40px;
            color: rgb(0, 0, 0);
            background-color: rgb(239, 239, 245);
            padding-left: 8px;
        }

        #addCategoryPopup {
            display: none;
            width: 700px;
            height: 230px;
            position: fixed;
            bottom: 300px;
            right: 400px;
            border: 1px solid #611212;
            z-index: 9;
            background-color: rgb(238, 238, 238);
        }

        #labelCategory {
            width: auto;
            height: 55px;
            border: 1px solid rgba(187, 173, 173, 0.918);
        }

        #manageCategoryPopup {
            display: none;
            width: 700px;
            height: 330px;
            position: fixed;
            bottom: 300px;
            right: 400px;
            border: 1px solid #611212;
            z-index: 9;
            background-color: rgb(238, 238, 238);
        }

        #snapShot {
            margin-top: 20px;
            margin-left: 50px;
            border: 2px solid black;
            width: 150px;
            height: 40px;
            color: aliceblue;
            background-color: rgb(100, 100, 183);
            padding-left: 8px;
        }
    </style>


</head>





<body>

    <!-- <input type='button' value='Export to XLSX' id="downloadExcelData"> -->
    <input type="button" id="addCategory" value="Add Category">
    <input type="button" id="manageCategory" value="Manage Category">

    <!-- <div id="popUp"></div> -->

    <select name="category" id="category">
        <option value="">- -Select Category- -</option>
        <option value="Project">Project</option>
        <option value="Software">Software</option>
        <option value="Hardware">Hardware</option>
    </select>

    <input type="button" id="insert" value="+Insert">



    <div id="addCategoryPopup">
        <div id="labelCategory1">
            <h1 style="margin: 5px 2px 2px 90px;">Add New Category</h1>
            <div style="width: 50px;height:50px; margin-top:-43px;margin-left: 620px;"><img
                    src="https://cdn-icons-png.flaticon.com/512/58/58253.png" id="crossImage1" alt="" width="40"
                    height="50"></div>
        </div>
        <p style="margin-top: 35px;">
            <label for="CategoryName1" style="font-size: 25px; margin-left: 30px;"><b>Category Name</b></label>
            <input type="text" id="addingCategoryField"
                style="width: 400px; height:30px; margin-left: 30px; padding-left: 20px;" placeholder="" name="email"
                required>
        </p>
        <input type="button" id="save1" value="Save"
            style="width: 80px; height: 35px; margin-left:600px; margin-top: 30px;">
    </div>

    <div id="manageCategoryPopup">
        <div id="labelCategory2">
            <h1 style="margin: 5px 2px 2px 90px;">Manage Category</h1>
            <div style="width: 50px;height:50px; margin-top:-43px;margin-left: 620px;"><img
                    src="https://cdn-icons-png.flaticon.com/512/58/58253.png" id="crossImage2" alt="" width="40"
                    height="50"></div>
        </div>
        <p id="addingNewCategoryByField" style="margin-top: 35px;">
            <p id="addingOption">Please select your category to apply:</p>
            <input type="radio" id="html" name="fav_language" value="HTML">
            <label for="html">HTML</label><br>
            <input type="radio" id="css" name="fav_language" value="CSS">
            <label for="css">CSS</label><br>
            <input type="radio" id="javascript" name="fav_language" value="JavaScript">
            <label for="javascript">JavaScript</label>
        </p>
        <input type="button" id="apply" value="apply" style="width: 80px; height: 35px; margin-left:600px; ">
    </div>


    <input type="button" id="snapShot" value="Baseline / Snapshot" style="display: block;">


    <div id="spreadsheet"></div>


    <!-- <script src="JSpreadSheet.js" type="module"></script> -->

    <script>
        // const downloadExcel = function () {
        //     jspreadsheet.render(document.getElementById('spreadsheet'), {
        //         filename: 'file.xlsx',
        //     });
        // }


        // jspreadsheet.setLicense("eyJhbGciOiJIUzUxMiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjoiMjg3NCIsInVzZXJfc2lnbmF0dXJlIjoiZWE5NzhhYzktNmUzNi00MWYwLTkxNGMtZDBhMzdiMzNiYWU2In0.ZEdadFwXXubQDM64MVhDtC638PuuhZev-K8pozyJMmLQRBBFqbTLh3bs3mUgOgejcZ6KQZu6wbIuFqZe39yC-Q");

        var data1 = [
            ['Total', 'Baseline', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            ['Total', 'Forecast', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            ['Total', 'Actual', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            ['Total', 'Variance', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ];


        // for(let i=0;i<data1.length;i++){
        //     data1[i][0]="Total";
        //     for(let j=0;j<data1.length[0];j++){
        //         data1[i+1][j+4]=newData[i+1][j+4];
        //     }
        // }
        // data1[1][0]="Baseline";
        // data1[2][0]="Forecast";
        // data1[3][0]="Actual";
        // data1[4][0]="Variance";


        const d = new Date();
        let year = d.getFullYear();


        function tableData() {
            var table = jspreadsheet(document.getElementById('spreadsheet'),
                {
                    //data: data1,
                    // columnDrag: true,
                    columns: getCostGridColumns(year),
                    // mergeCells: {
                    //     A1: [1, 4]
                    // },

                });

            table.hideIndex();

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


            let columnData = document.getElementById("abc");
            console.log(columnData)
            columnData.append(previous);
            columnData.append(year);
            columnData.append(next);
            var br = document.createElement("br");
            columnData.append(br);
            columnData.append('FY TOTAL');


            document.getElementById("insert").addEventListener("click", function () {
                let selectionData = document.getElementById("category");
                let val = selectionData.value;
                // if(val==null){
                //     document.getElementById("selectCat").style.display = "block";
                // }
                selectionData.remove(selectionData.selectedIndex);
                var mergeIndex = data1.length + 1;


                if (val !== "") {
                    var newData = [
                        [val, 'Baseline', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                        [val, 'Forecast', 0, `=SUM(E${mergeIndex + 1}:P${mergeIndex + 1})`, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                        [val, 'Actual', 0, `=SUM(E${mergeIndex + 2}:P${mergeIndex + 2})`, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                        [val, 'Variance', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                    ]

                    // console.log(table)
                    console.log("merge", `A${mergeIndex}`);

                    data1 = data1.concat(newData);
                    console.log(newData)
                    table.setData(data1);
                    // for(let i=data.length-4;i<data1.length-4;i++){
                    //     data1[2][i]+=`=SUM(E${i+1}:E${i+1})`;
                    //     data1[3][i]+=table.getData()[mergeIndex+2][i];
                    //     console.log(table.getData()[mergeIndex+1][i]);
                    // }
                    // table.setData(data1);
                    if (mergeIndex === 5) {
                        table.setMerge("A1", 1, 4);
                    }

                    table.setMerge(`A${mergeIndex}`, 1, 4);
                }
                //console.log(data1);

                data1[1][3] = `=SUM(E${mergeIndex + 1}:P${mergeIndex + 1})`;
            })
        }

        tableData();

        function updateYear(year, previous, next) {
            let columnData = document.getElementById("abc");
            columnData.innerText = '';
            columnData.append(previous);

            columnData.append(year);
            columnData.append(next);
            var br = document.createElement("br");
            columnData.append(br);
            columnData.append('FY TOTAL');

            let jan = document.getElementById("JAN");
            jan.innerText = "JAN " + year;
            let feb = document.getElementById("FEB");
            feb.innerText = "FEB " + year;
            let mar = document.getElementById("MAR");
            mar.innerText = "MAR " + year;
            let apr = document.getElementById("APR");
            apr.innerText = "APR " + year;
            let may = document.getElementById("MAY");
            may.innerText = "MAY " + year;
            let jun = document.getElementById("JUN");
            jun.innerText = "JUN " + year;
            let jul = document.getElementById("JUL");
            jul.innerText = "JUL " + year;
            let aug = document.getElementById("AUG");
            aug.innerText = "AUG " + year;
            let sep = document.getElementById("SEP");
            sep.innerText = "SEP " + year;
            let oct = document.getElementById("OCT");
            oct.innerText = "OCT " + year;
            let nov = document.getElementById("NOV");
            nov.innerText = "NOV " + year;
            let dec = document.getElementById("DEC");
            dec.innerText = "DEC " + year;
        }

        let columnData = document.getElementById("lifetime");
        var br = document.createElement("br");
        columnData.append(br);
        columnData.append('TOTAL');



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
                    width: 120,
                    id: 'lifetime',
                    title: 'LIFE TIME',
                    // align: 'right',
                    wordWrap: true,
                    readOnly: true,
                    // //	mask:'#,##,00', decimal:',',
                    // font: 'Verdana,Arial,sans-serif'

                },
                {
                    type: 'text',
                    width: '120',
                    title: ' ',
                    id: 'abc',
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
                        width: '80',
                        id: months[i],
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

        document.getElementById('save1').addEventListener('click', saveCategory);

        document.getElementById('crossImage1').addEventListener('click', closeForm1);

        document.getElementById('manageCategory').addEventListener('click', openForm2);

        document.getElementById('apply').addEventListener('click', newCategoriesAdded);

        document.getElementById('crossImage2').addEventListener('click', closeForm2);

        // copying forecast data to baseline
        document.getElementById('snapShot').addEventListener('click', screenShot);




        // document.getElementById('downloadExcelData').addEventListener('click', downloadExcel);


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
            console.log("print");
            document.getElementById("manageCategoryPopup").style.display = "none";
        }

        function screenShot() {

        }

        function saveCategory() {
            document.getElementById("addCategoryPopup").style.display = "none";
            var br = document.createElement("br");
            document.getElementById("addingNewCategoryByField").append(br);
            let newCategoryVal = document.getElementById("addingCategoryField").value;
            var radioInput = document.createElement('input');
            radioInput.setAttribute('type', 'radio');
            radioInput.setAttribute('name', name);
            document.getElementById("addingNewCategoryByField").appendChild(radioInput);
            var elem2 = document.createElement("label");
            elem2.innerHTML = newCategoryVal;
            document.getElementById("addingNewCategoryByField").appendChild(elem2);

        }

        function newCategoriesAdded() {

        }

        // function newCategoriesAdded() {
        //     document.getElementById("manageCategoryPopup").style.display = "none";
        //     console.log(document.getElementById("addingOption").value);
        // }

        // Add-on for Jspreasheet
        // jspreadsheet.setExtensions({ render });

    </script>


</body>

</html>