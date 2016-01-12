/*!
 * index.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */

var init_100 = function(){
  console.log("check_evn_100");
  setInterval(function() {
    
    var max_count = 5;
    var count = vm.run_msg.length;
    if( count > max_count ){
      console.log(count-max_count);
      vm.run_msg = _.rest(vm.run_msg);
    }
  }, 3000);
}

var check_evn_101 = function(){
  console.log("check_evn_101");

  // 程序运行所必须的库和配置文件
  var envlist = [];
  envlist.push("config.json");
  envlist.push("node_modules/node-xlsx/");
  envlist.push("node_modules/underscore/underscore-min.js");

  vm.step = 102;
  return "";
};

var check_status_102 = function(){
  console.log("check_status");
  var run_flag = true;

  // 程序运行所必须的数据源
  var srcFilelist = [];
  srcFilelist.push("11月销售订单明细.XLSX");
  srcFilelist.push("基础价格表.XLSX");
  srcFilelist.push("价格变动表.XLSX");
  srcFilelist.push("价保返利明细表.XLSX");
  srcFilelist.push("客户地市对应表.XLSX");
  srcFilelist.push("渠道经理对应表.XLSX");
  srcFilelist.push("产品经理对应表.XLSX");
  //srcFilelist.push("abc.XLSX");

  for( var i=0;i<srcFilelist.length;i++ ){
    var filename = srcFilelist[i];
    if(fs.existsSync("data/" + filename)){
      put_msg( "必要文件：" + filename + " 存在，程序可以继续工作。");
      run_flag = true;
    }else{
      put_err_msg( "必要文件： " + filename + " 不存在，程序无法继续工作。");
      run_flag = false;
      break;
    }
  }
  console.log(vm.run_msg);

  if( run_flag ){
    vm.step = 103;
  }else{
    vm.step = 101;
  }

  return "";
};

var check_src_103 = function(){
  console.log("check_src");
  vm.step = 104;
  return "";
};

var load_order_detail_104 = function(){
  console.log("load_order_detail");
  put_msg( "数据较多，载入约需15秒。请耐心等待。");
  vm.step = 104.5;

  var obj = null;
  setTimeout(function() {
    obj = xlsx.parse('data/11月销售订单明细(全).xlsx'); // 读入xlsx文件
    // head = obj[0].data[0];
    ORDER_DETAIL = obj[0].data;   // 
    //console.log(obj[0].name);
    put_msg( "销售订单明细数据读入成功。");
    vm.step = 105;
  }, 100);
  

  
  return "";
};

var copy_order_detail_105 = function(){
  console.log("copy_order_detail");

  
  var must_col_idx = [];
  var must_col_title = [];

  must_col_idx.push(0);
  must_col_title.push("实际交货数量");
  must_col_idx.push(2);
  must_col_title.push("交货完成状态");
  must_col_idx.push(3);
  must_col_title.push("销售价格");
  must_col_idx.push(4);
  must_col_title.push("销售金额");
  must_col_idx.push(5);
  must_col_title.push("订单类型");
  must_col_idx.push(6);
  must_col_title.push("渠道");

  must_col_idx.push(8);
  must_col_title.push("客户");
  must_col_idx.push(9);
  must_col_title.push("客户名称");

  must_col_idx.push(12);
  must_col_title.push("销售数量");

  must_col_idx.push(30);
  must_col_title.push("物料号");
  must_col_idx.push(31);
  must_col_title.push("物料组");
  must_col_idx.push(32);
  must_col_title.push("物料组描述");

  must_col_idx.push(47);
  must_col_title.push("实际交货日期");
  must_col_idx.push(48);
  must_col_title.push("城市");

  must_col_idx.push(54);
  must_col_title.push("客户参考号");

  must_col_idx.push(64);
  must_col_title.push("创建日期");


  var must_col_sheet = [];
  for( var i=0; i < ORDER_DETAIL.length; i++ ){
    must_col_content = pick_from_array(must_col_idx, ORDER_DETAIL[i]);
    must_col_sheet.push(must_col_content);
    // console.log(must_col_content);
    console.log(i);
  }

  var buffer = xlsx.build([{name: "销售订单明细", data: must_col_sheet}]);
  fs.writeFileSync( "data/销售订单明细精简版.xlsx", buffer);

  // 开发时使用的工具代码
  // var title = ORDER_DETAIL[0];
  // for( var i=0; i < title.length; i++ ){
  //   console.log("" + i + ":" + title[i]);
  // }


  vm.step = 106;
  return "";
};

var pick_from_array = function(index_array, src_array){

  var dest_array = [];
  for(var i=0; i<index_array.length; i++){
    var idx = index_array[i];
    dest_array.push(src_array[idx]);
  }
  return dest_array;
}