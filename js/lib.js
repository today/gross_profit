/*!
 * lib.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */


var Msg = {
  createNew: function(){
    var msg = {};
    msg.msg_list = [];

    msg.put = function(a_msg){ 
      msg.msg_list.push("log " + msg.msg_list.length + ":" + a_msg); 
    };

    return msg;
  }
};

function trim_array_element( a_array ){
  for( var i=0; i<a_array.length; i++){
    var temp = a_array[i];
    if( temp ){
      if( (typeof temp) == "string" ){
        a_array[i] = temp.trim();
      }else{
        a_array[i] = temp.toString().trim();
      }
    }
  }
}

function isblank(strA){
  if(strA){
    if( "string" === typeof(strA) ){
      if( "" === strA.trim()){
        return true;
      }else{
        return false;
      }
    }else{
      return false;
    }
  }else{
    return true;
  }
}