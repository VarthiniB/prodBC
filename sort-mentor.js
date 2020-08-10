
var Excel = require('exceljs');

var workbook = new Excel.Workbook();

workbook.xlsx.readFile('final.xlsx')
    .then(function() {
     
        var worksheet5 = workbook.getWorksheet('final');
        var worksheet6 = workbook.getWorksheet('sort-mentor');

        var top=[];
       for(let i=2;i<18;i++){
        var arr=[];
           for(let j=2;j<23;j++){

            let flagrow2 = worksheet5.getRow(i);
            let flag1cell = flagrow2.getCell(j);
            let flagrow3 = worksheet5.getRow(1);
            let flag1cellname = flagrow3.getCell(j);
            var person = {firstName:flag1cellname, score:flag1cell};
              arr.push(person);
       
           }
          top.push(arr);
       }
      for (let k = 0; k < top.length; k++) {
        // console.log( top[k].toString());
        for (let p = 0; p < top[k].length; p++) {
            // console.log( top[k].toString());
            top[k].sort(function(a, b){return b.score - a.score});
            // console.log("------------"+top[k][p].firstName+"---"+top[k][p].score);             
          }
        // console.log("=======================");          
      }
      let mm=2
      let rowcounter = 1;
      for (let k = 0; k < top.length; k++) {   
          let flagrow2 = worksheet5.getRow(mm++);
          let flag1cell = flagrow2.getCell(1);
          console.log("mentor :"+flag1cell);
          row = worksheet6.getRow(rowcounter++);
          row.getCell(1).value = flag1cell.text;
          row.commit();
        for (let p = 0; p < top[k].length; p++) {          
          console.log("                     "+top[k][p].firstName+"   "+top[k][p].score);   
          row = worksheet6.getRow(rowcounter++);
          row.getCell(2).value = top[k][p].firstName.text;
          row.getCell(3).value = top[k][p].score.text;
          row.commit();
        }
      }
      return workbook.xlsx.writeFile('sort-mentor.xlsx');
    })
     