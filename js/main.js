//function getJSONData(path, callback) {
//    var httpRequest = new XMLHttpRequest();
//    httpRequest.onreadystatechange = function() {
//        if (httpRequest.readyState === 4) {
//            if (httpRequest.status === 200) {
//                var data = JSON.parse(httpRequest.responseText);
//                if (callback) callback(data);
//            }
//        }
//    };
//    httpRequest.open('GET', path);
//    httpRequest.send(); 
//}
//
//
//getJSONData('/json.php', function(data){
//    console.log(data);
//    console.log(data.a);
//});




// node.js server
function vote() {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', 'vote', true);
  xhr.onreadystatechange = function() {
    if (xhr.readyState != 4) return;

    if (xhr.status != 200) {
      // обработать ошибку
      alert('Ошибка ' + xhr.status + ': ' + xhr.statusText);
      return;
    }

    // обработать результат
    console.log( JSON.parse(xhr.response) );
  }

  xhr.send(null);

}
vote();