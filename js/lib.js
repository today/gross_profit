/*!
 * lib.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */
var _ = require('underscore');

var Msg = {
  createNew: function(){
    var msg = {};
    msg.msg_list = [];
    msg.msg_count = 0;

    msg.put = function(a_msg){ 
      msg.msg_list.push("log " + msg.msg_list.length + ":" + a_msg); 
      msg.msg_count++;
    };

    return msg;
  }
};

/*
 从二维数组中删除指定的列 
*/
function del_col_from_array(a_array, col_indexs){
  //console.log(a_array);
  var src_col_index = _.range(a_array[0].length);
  var select_index = _.difference(src_col_index, col_indexs);
  return select_col_from_array(a_array, select_index);
}

/*
 从二维数组中取出指定的列 
*/
function select_col_from_array(a_array, col_indexs){
  var dest = [];
  var temp = [];
  for(var i=0; i<a_array.length; i++){
    temp = [];
    for(var j=0; j<col_indexs.length; j++){
      temp.push(a_array[i][col_indexs[j]]);
    }
    dest.push(temp);
  }
  return dest;
}

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