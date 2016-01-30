/*!
 * lib.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */
var _ = require('underscore');
var fs = require('fs');
var path = require('path');

var Msg = {
  createNew: function(){
    var msg = {};
    msg.msg_list = [];
    msg.msg_count = 0;

    msg.put = function(a_msg){ 
      //msg.msg_list.push("log " + msg.msg_list.length + ":" + a_msg); 
      msg.msg_list.push( a_msg ); 
      msg.msg_count++;
    };

    msg.download = function(){
      var evt = document.createEvent("HTMLEvents");
      evt.initEvent("click", false, false);//initEvent 不加后两个参数在FF下会报错
      
      var a_txt = [];
      msg.msg_list.forEach(function(e){  
           a_txt.push(e);
           a_txt.push("\n");  
      });

      var blob = new Blob(a_txt);

      var aLink = document.createElement('a');
      aLink.download = "错误数据记录.txt";
      aLink.href = URL.createObjectURL(blob, { "type" : "text/xml" });
      aLink.dispatchEvent(evt);
    }

    return msg;
  }
};


// 根据关键词，寻找文件名匹配的目标文件。
function find_src_file( base_dir, files_flag ){
  var dest_files = [];

  if(fs.existsSync(base_dir)) {
    console.log('base_dir 存在');

    var all_file = fs.readdirSync(base_dir);

    for(var i=0; i<files_flag.length; i++){
      var key=files_flag[i];
      var file_name = undefined;

      for(var j=0; j<all_file.length; j++){
        if( all_file[j].indexOf(key) > -1 ){
          var file_name = all_file[j];
        }
      }
      dest_files[key] = file_name;
    }
  } else {
    console.log('base_dir 不存在');
    return null;
  }
  console.log(dest_files);
  return dest_files;
}

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