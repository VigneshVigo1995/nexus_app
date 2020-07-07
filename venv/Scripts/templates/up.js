    function openDialog() {
         document.getElementById('fileid').click();
         }
    function Handlechange()
         {
           var fileinput = document.getElementById("browse");
           var textinput = document.getElementById("filename");
           textinput.value = fileinput.value;
         }

    function export () {

    var textFile = $('.postingFile textarea').val();

    var element = document.createElement('a');

    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(textFile));

    element.setAttribute('download', 'filename.csv');

    element.style.display = 'none';

  if (typeof element.download != "undefined") {

      //browser has support - process the download

      document.body.appendChild(element);

      element.click();

      document.body.removeChild(element);

    }

    else {

      //browser does not support - alert the user

      alert('This functionality is not supported by the current browser, recommend trying with Google Chrome instead.  (http://caniuse.com/#feat=download)');

    } //end of if...else...block

; //end of $('.downloadCSVFileButton').click(function () {
