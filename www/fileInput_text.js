$(document).ready(function(){
  $('#file_progress').on("DOMSubtreeModified",function(){

    var target = $('#file_progress').children()[0];
    if(target.innerHTML === "Upload complete"){
        console.log('Change')
        target.innerHTML = '';      
    }

  });
});