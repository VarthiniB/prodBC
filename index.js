
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('env.xlsx')
    .then(function() {
        var worksheet4 = workbook.getWorksheet(4);
        var worksheet5 = workbook.getWorksheet('final');
        console.log("==1=="+worksheet4.name);
        console.log("==2=="+worksheet5.name);
      //  var row = worksheet.getRow(5);
      //  row.getCell(1).value = 50000; // A5's value set to 5
    //  var row52 = worksheet5.get(2);
    //  var row42 = worksheet4.get(28);
    var k=2;
       for(let i=28;i<44;i++){
        let q = k++;
        let w = 2;
           for(let j=4;j<25;j++){
               //console.log("===i"+i+"====j"+j);
              let flagrow1 = worksheet4.getRow(i);
               let flagrow2 = worksheet4.getRow(j);             
                let wfor = w++;               
               var rowS5 = worksheet5.getRow(q);
               rowS5.getCell(wfor).value = 0;


               let flag1cell = flagrow1.getCell(2);
               let flag2cell = flagrow2.getCell(2);
             
               if((flag1cell - flag2cell) > 2){
                   rowS5.getCell(wfor).value = flag1cell - flag2cell;
               }
               else{
                   rowS5.getCell(wfor).value = 0;
               } 
             
             let flag1cellDeg1 = flagrow1.getCell(4);
               let flag2cellDeg2 = flagrow2.getCell(4);

               if((flag1cellDeg1 - flag2cellDeg2) < 0){
                rowS5.getCell(wfor).value =  rowS5.getCell(wfor).value - 1;
               }
               else{
                    if((flag1cellDeg1 - flag2cellDeg2) == 0 ){
                        rowS5.getCell(wfor).value =  rowS5.getCell(wfor).value + 1;
                    }
                    else{
                        rowS5.getCell(wfor).value =  rowS5.getCell(wfor).value + 2;
                    }
               } 
               
               let flag1cellIndus1 = flagrow1.getCell(9)+ '';
               let flag1cellIndus2 = flagrow2.getCell(9)+ '';
                const Indus1 = flag1cellIndus1.split(';');
                const Indus2 = flag1cellIndus2.split(';');
               let zI=0;
               for(let x=0;x<Indus1.length;x++){
                   for(let y=0;y<Indus2.length;y++){
                        if(Indus2[y] == Indus1[x] && Indus2[y] != "" && Indus1[x] != ""){
                            zI++;
                        }
                   }
               }               
                rowS5.getCell(wfor).value =  rowS5.getCell(wfor).value + zI++;

               let flag1cellEnv1 = flagrow1.getCell(10)+ '';
               let flag1cellEnv2 = flagrow2.getCell(10)+ '';
                const Env1 = flag1cellEnv1.split(';');
                const Env2 = flag1cellEnv2.split(';');
               let zE=0;
               for(let x=0;x<Env1.length;x++){
                   for(let y=0;y<Env2.length;y++){
                        if(Env2[y] == Env1[x] && Env2[y] !="" && Env1[x] != ""){
                            zE++;
                        }
                   }
               }
               rowS5.getCell(wfor).value =  rowS5.getCell(wfor).value + zE++;
              
               let flag1cellLoc1 = flagrow1.getCell(31)+ "";
               let flag2cellLoc2 = flagrow2.getCell(31)+ "";
                  // console.log("=============="+flag1cellLoc1);
                  // console.log("=============="+flag2cellLoc2);
               const location1 = flag1cellLoc1.split(';');
               const location2 = flag2cellLoc2.split(';');
              let z=0;
              for(let x=0;x<location1.length;x++){
                  for(let y=0;y<location2.length;y++){
                       if(location2[y] == location1[x] && location2[y] != "" && location1[x] != ""){
                           z++;
                       }
                  }

              }
              //console.log("===z="+z);
              if(z==0){
               rowS5.getCell(wfor).value =  0;
              }
              //only for seperate sheet not for final
        /*      else{
               rowS5.getCell(wfor).value =  z;
              }*/

              let flag1cellTime1 = flagrow1.getCell(32)+ '';
              let flag2cellTime2 = flagrow2.getCell(32)+ '';

               const Time1 = flag1cellTime1.split(';');
               const Time2 = flag2cellTime2.split(';');
              let zT=0;
              for(let x=0;x<Time1.length;x++){
                  for(let y=0;y<Time2.length;y++){
                       if(Time2[y] == Time1[x] && Time2[y] != "" && Time1[x] !=""){
                           zT++;
                       }
                  }

              }
              if(zT==0){
               rowS5.getCell(wfor).value =  0;
              }
              //only for seperate sheet not for final
           /*   else
              {
                   rowS5.getCell(wfor).value =  zT;
              }*/

               rowS5.commit();
           }
       }
    
    
        return workbook.xlsx.writeFile('final.xlsx');
    })