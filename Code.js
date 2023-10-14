function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Google Classroom Tools')
      .addItem('Listar Clases','getClassrooms')
      .addItem('Detalles Clases','getClasswork')
      .addItem('Auto-update listing every night', 'createTimeDrivenTriggers')
      .addToUi();
      
}

function getClassrooms(){

 var ssP=SpreadsheetApp.getActive();
  var sh=ssP.getSheetByName("RESUMEN");
  var cvePlantel=sh.getRange(2,10).getValue();
  //Logger.log(cvePlantel);
  sh.getRange(3,11).setValue("Clases iniciado "+cvePlantel );

  var fileArray = [["c.name","c.updateTime","c.id","emailAddress","alumnos","semestre","plantel"]];
var pageToken, page;
do{
   var optionalArgs = {"courseStates":"ACTIVE",pageToken: pageToken};  
   
      page = Classroom.Courses.list(optionalArgs);
      var allCourses = page.courses;
      if (!allCourses){Logger.log( page.courses);
      }
      var ssData=allCourses.map(c =>{
        return[c.name,c.updateTime,c.id,c.ownerId];
      })
      var str="2021";
      var studiantes=0;
      var ssDataf=ssData.filter(function(row){
                                return row[1].toString().indexOf(str) > -1;
                          });
                      
      var newssData=ssDataf.map(function(row){return[row[0],row[1],row[2],row[3],row[2],row[0],row[0]]
                                  });
var nssDataF=newssData.filter(
  function(row){ return row[5]==1|| row[5]==3|| row[5]==5; });


function plantel(nombre){
  var nombrel=nombre;
  if (nombre.length<20){nombrel=nombre +" materia";}
  var primer;
  var segundo;
  var tercero;
  var nuevonombre;
  if (nombrel.indexOf("-")>-1){ 
primer=nombrel.search("-")+1;
 nuevonombre=nombrel.substring(primer,nombrel.length-primer);
if (nuevonombre.search("-")+1>0){
 
segundo=nuevonombre.search("-")+1;
 tercero= "'"+nuevonombre.substring(7,10);
                    return tercero;
}



  }else {return "";}


}

function alumnos(idclase){
var vargs = {"pageSize": 40};

const roster = []; options = {pageSize: 50};
do {
  // Get the next page of students for this course.
  var search = Classroom.Courses.Students.list(idclase, options);

  // Add this page's students to the local collection of students.
  // (Could do something else with them now, too.)
  if (search.students)
    Array.prototype.push.apply(roster, search.students);

  // Update the page for the request
  options.pageToken = search.nextPageToken;
} while (options.pageToken);



  //"courseId":
  try{return roster.length;//Classroom.Courses.Students.list(idclase,vargs).students.length;


  }
  catch (err){return 0;}
}


      Array.prototype.push.apply(fileArray, nssDataF);
        pageToken = page.nextPageToken;

}
while (pageToken);
var ss=SpreadsheetApp.getActiveSpreadsheet();
var trn=ss.getSheetByName("Classes");
trn.getRange(1,1,fileArray.length,fileArray[5].length).setValues(fileArray);
var n047array=fileArray.filter(
  function(row){ return row[6]=="'"+cvePlantel; });
  var datosn=n047array.map(function(row){
    return[row[0],row[1],row[2],row[3],row[4],row[5],row[6],alumnos(row[2]),getTeacher(row[3])];
    });
  var trn047=ss.getSheetByName(cvePlantel);
trn047.getRange(2,1,datosn.length,datosn[5].length).setValues(datosn);
var rangleclr=trn047.getRange(2,10,datosn.length,4);
rangleclr.clearContent();

var rclear2=trn047.getRange(2,18,datosn.length,2);
rclear2.clearContent();
sh.getRange(2,11).setValue("Clases Completado "+cvePlantel);
}

function getStudents(){
  var ss=SpreadsheetApp.getActive();
  var sh=ss.getSheetByName("264");
  var datos=sh.getRange(2,3,sh.getLastRow()-1,1).getValues();
  //var idprofe="107801114177388388476";
  var datosn=datos.map(function(row){
    return[alumnos(row[0])];
    });
    
    sh.getRange(2,17,datosn.length,datosn[0].length).setValues(datosn);

}

function alumnos(idclase){

var roster2 = [];
var optionalArgs2 = {"courseId": idclase};
var ptoken2;
var options2 = {"courseId": idclase,pageSize:1,pageToken : ptoken2 };
 var correo;
do {
  // Get the next page of students for this course.


  var search = Classroom.Invitations.list(options2);

  // Add this page's students to the local collection of students.
  // (Could do something else with them now, too.)
  if (search.invitations)
    //Array.prototype.push.apply(roster, search.userId);
roster2=search.invitations.map(d =>{
        return[getTeacher(d.userId)];
      });
      if(roster2){if (roster2!=="undefined"){correo=roster2 + ","+correo;}
        }
      
      
  // Update the page for the request
 // Logger.log(search.nextPageToken);
  options2.pageToken = search.nextPageToken;
  ptoken2=search.nextPageToken;
 // Logger.log(ptoken2);
} while (options2.pageToken);

//Logger.log(correo);
return correo;
}

function getTeacher(idprofesor) {

  var cache = CacheService.getScriptCache();
  var cached = cache.get(idprofesor);
  if (cached != null) {
    return cached;
  }
  try{
  var result = Classroom.UserProfiles.get(idprofesor).emailAddress; // takes 20 seconds
  //var contents = result.getContentText();
  cache.put(idprofesor, result, 1500); // cache for 25 minutes
  return result;}
  catch(err){return }

}

function getClasswork()
{
   var ssP=SpreadsheetApp.getActive();
  var shr=ssP.getSheetByName("RESUMEN");
  var cvePlantel=shr.getRange(2,10).getValue();
  //Logger.log(cvePlantel);
   shr.getRange(3,11).setValue("Detalles Iniciado "+cvePlantel);
  
  var ss=SpreadsheetApp.getActive();
  var sh=ss.getSheetByName(cvePlantel);
  var datos=sh.getRange(2,3,sh.getLastRow()-1,2).getValues();
  var datosn=datos.map(function(row){
    return[getProfe(row[0]),getInvites(row[0]),getCW(row[0])+getAsignP(row[0]),getCWd(row[0])+ getAsignD(row[0])];
    });
function getProfe(idc){
          try{return Classroom.Courses.Teachers.list(idc).teachers[0].profile.emailAddress;    

          }
          catch (err){return 0;
                   // Logger.log(err);
                   }
                      



}



      function getCW(idc){
              var param= {"courseWorkMaterialStates":"PUBLISHED"};  
              try{return Classroom.Courses.CourseWorkMaterials.list(idc,param).courseWorkMaterial.length;}
              catch (err){return 0;
                   // Logger.log(err);
                   }
                      }

      function getAsignP(idc){
            var param= {"CourseWorkStates":"PUBLISHED"};  
            try{ return Classroom.Courses.CourseWork.list(idc,param).courseWork.length;}
            catch(err){return 0;//Logger.log(err);
            }
      }
      function getAsignD(idc){
            var param= {"CourseWorkStates":"DRAFT"};  
            try{ return Classroom.Courses.CourseWork.list(idc,param).courseWork.length;}
            catch(err){return 0;//Logger.log(err);
            }
      }




function getCWd(idc){
              var param= {"courseWorkMaterialStates":"DRAFT"};  
              try{return Classroom.Courses.CourseWorkMaterials.list(idc,param).courseWorkMaterial.length;}
              catch (err){return 0;
                   // Logger.log(err);
                   }
                      }

function getInvites(idc){
          var optionalArgs2 = {"courseId": idc};
          try{return Classroom.Invitations.list(optionalArgs2).invitations.length;
          }
          catch(err){return 0;
          //Logger.log(err);
          }

}

    sh.getRange(2,10,datosn.length,datosn[0].length).setValues(datosn);
    sh.setFrozenRows(1);
    sh.sort(1);
     shr.getRange(3,11).setValue("Detalles Completado" +cvePlantel);
}

function createTimeDrivenTriggers() {
  // Trigger every night at 11 pm
ScriptApp.newTrigger('getClassrooms').timeBased().everyDays(1).atHour(23).create();            
}