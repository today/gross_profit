/*!
 * index.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */

var init_100 = function(){
  console.log("check_evn_100");
  MSG.put("系统启动。");

  setInterval(function() {
    
    var max_count = 15;
    var count = vm.run_msg.length;
    if( count > max_count ){
      console.log(count-max_count);
      //vm.run_msg = _.rest(vm.run_msg);
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
  srcFilelist.push("11月销售订单明细(全).XLSX");
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
      MSG.put( "必要文件：" + filename + " 存在，程序可以继续工作。");
      run_flag = true;
    }else{
      ERR_MSG.put( "必要文件： " + filename + " 不存在，程序无法继续工作。");
      run_flag = false;
      break;
    }
  }

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
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");
  vm.step = 104.5;

  var obj = null;
  setTimeout(function() {
    obj = xlsx.parse('data/2sheet.xlsx'); // 读入xlsx文件
    //obj = xlsx.parse('data/11月销售订单明细(全).xlsx'); // 读入xlsx文件

    // head = obj[0].data[0];
    ORDER_DETAIL = obj[0].data;   // 
    //console.log(obj[0].name);
    MSG.put( "销售订单明细数据读入成功。");
    vm.step = 105;
  }, 100);
  

  
  return "";
};


/* 
毛利 = （销售收入-成本单价*交货数量）/1.17-下游价返-促销费
*/
var copy_order_detail_105 = function(){
  console.log("copy_order_detail");

  var must_col_title = [];
  must_col_title.push(make_title("实际交货数量"));
  must_col_title.push(make_title("交货完成状态"));
  must_col_title.push(make_title("销售价格"));
  must_col_title.push(make_title("销售金额"));
  must_col_title.push(make_title("订单类型"));
  must_col_title.push(make_title("渠道"));
  must_col_title.push(make_title("客户"));
  must_col_title.push(make_title("客户名称"));
  must_col_title.push(make_title("销售数量"));
  must_col_title.push(make_title("物料号"));
  must_col_title.push(make_title("物料组"));
  must_col_title.push(make_title("物料组描述"));
  must_col_title.push(make_title("实际交货日期"));
  must_col_title.push(make_title("城市"));
  must_col_title.push(make_title("客户参考号"));
  must_col_title.push(make_title("创建日期"));
  // 以下数据，要从其他文件中获得
  must_col_title.push(make_title("成本单价"));
  must_col_title.push(make_title("下游价返"));
  must_col_title.push(make_title("促销费"));
  must_col_title.push(make_title("利润"));
  must_col_title.push(make_title("利润率"));

  
  var title_array = ORDER_DETAIL[0];
  // 把附加字段添加到 title 行中。
  title_array.push("成本单价");
  title_array.push("下游价返");
  title_array.push("促销费");
  title_array.push("利润");
  title_array.push("利润率");

  // 根据title，查询出 index。
  for(var i=0; i<title_array.length; i++ ){
    for(var j=0; j<must_col_title.length; j++){
      t1 = title_array[i];
      obj_title = must_col_title[j];
      t2 = obj_title.title;
      if( t1 == t2 ){
        obj_title.index = i;
      }
    }
  }

  // 获取所有的 index 值。
  var must_col_idx = _.pluck(must_col_title, 'index');
  console.log(must_col_idx);


  var must_col_content = [];

  var must_col_sheet = [];
  for( var i=0; i < ORDER_DETAIL.length; i++ ){
    must_col_content = pick_from_array(must_col_idx, ORDER_DETAIL[i]);
    must_col_sheet.push(must_col_content);
    // console.log(must_col_content);
    //console.log(i);
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
    if( idx===-1 ){
      dest_array.push("");
    }
    else if( idx>=0 ){
      dest_array.push(src_array[idx]);
    }
    else{
      dest_array.push("n/a");
    }
  }
  return dest_array;
}

var make_title = function( s_title ){
  var obj_title = {};
  obj_title.index = -1;
  obj_title.title = s_title;
  
  return obj_title;
};













