var _ = require('underscore');
var fs = require('fs');

var path = "./data/201601";

var src_files_flag = [];
src_files_flag.push("销售订单明细");
src_files_flag.push("物料清单");
src_files_flag.push("SCM客户明细");
src_files_flag.push("内控成本价格变动");

var dest_files = [];


if(fs.existsSync(path)) {
  console.log('base_dir 存在');

  var all_file = fs.readdirSync(path);
  console.log(all_file);
  for(var i=0; i<src_files_flag.length; i++){
    var key=src_files_flag[i];
    var file_name = undefined;

    for(var j=0; j<all_file.length; j++){
      if( all_file[j].indexOf(key) > -1 ){
        var file_name = all_file[j];
      }
    }

    dest_files[key] = file_name;
  }
  console.log(dest_files);

} else {
  console.log('base_dir 不存在');
  return null;
}