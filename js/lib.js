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