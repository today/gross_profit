/*!
 * index.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */

var fs = require('fs');
var _ = require('underscore');
var xlsx = require("node-xlsx");

var ORDER_DETAIL = null;


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

var check_env_110 = function(){
  console.log("check_evn_101");

  // 程序运行所必须的库和配置文件
  var envlist = [];
  envlist.push("config.json");
  envlist.push("node_modules/node-xlsx/");
  envlist.push("node_modules/underscore/underscore-min.js");

  vm.step = 120;
  return "";
};

var check_status_120 = function(){
  console.log("check_status");
  var run_flag = true;

  // 程序运行所必须的数据源
  var srcFilelist = [];
  srcFilelist.push("data/201511/11月销售订单明细.XLSX");
  srcFilelist.push("data/201511/物料清单.XLSX");
  srcFilelist.push("data/201511/SCM客户明细.XLSX");
  srcFilelist.push("data/201511/2015.11月计提并使用返利.XLSX");
  srcFilelist.push("data/201511/2015.12促销品领用出库明细.XLSX");
  
  //srcFilelist.push("abc.XLSX");

  for( var i=0;i<srcFilelist.length;i++ ){
    var filename = srcFilelist[i];
    if(fs.existsSync("" + filename)){
      MSG.put( "必要文件：" + filename + " 存在，程序可以继续工作。");
      run_flag = true;
    }else{
      ERR_MSG.put( "必要文件： " + filename + " 不存在，程序无法继续工作。");
      run_flag = false;
      break;
    }
  }

  if( run_flag ){
    vm.step = 130;
  }else{
    vm.step = 110;
  }

  return "";
};

var check_src_130 = function(){
  console.log("check_src");
  vm.step = 140;
  return "";
};

var load_order_detail_140 = function(){
  console.log("load_order_detail_140");
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");
  vm.step = 140.5;

  var obj = null;
  setTimeout(function() {
    obj = xlsx.parse('data//201511/11月销售订单明细.xlsx'); // 读入xlsx文件
    //  obj = xlsx.parse('data/2sheet.xlsx'); // 读入xlsx文件

    // head = obj[0].data[0];
    ORDER_DETAIL = obj[0].data;   // 
    //console.log(obj[0].name);
    MSG.put( "销售订单明细数据读入成功。");
    vm.step = 150;
  }, 100);
  

  
  return "";
};


/* 
毛利 = （销售收入-成本单价*交货数量）/1.17-下游价返-促销费
*/
var copy_order_detail_150 = function(){
  console.log("copy_order_detail_150");
  
  var must_col_title = [];
  must_col_title.push(make_title("实际交货数量"));
  must_col_title.push(make_title("物料描述"));
  must_col_title.push(make_title("交货完成状态"));
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
  //must_col_title.push(make_title("下游价返"));
  //must_col_title.push(make_title("促销费"));
  must_col_title.push(make_title("利润"));
  must_col_title.push(make_title("利润率"));

  
  var title_array = ORDER_DETAIL[0];
  // 把附加字段添加到 title 行中。
  title_array.push("成本单价");
  //title_array.push("下游价返");
  //title_array.push("促销费");
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
  console.log(must_col_title);

  // 获取所有的 index 值。
  var must_col_idx = _.pluck(must_col_title, 'index');
  var must_col_title = _.pluck(must_col_title, 'title');
  console.log(must_col_idx);
  console.log(must_col_title);

  var must_col_content = null;
  var ORDER_DETAIL_SMALL = [];
  for( var i=0; i < ORDER_DETAIL.length; i++ ){
    var one_order = ORDER_DETAIL[i];
    if( !isblank(one_order[0]) ){
      must_col_content = pick_from_array(must_col_idx, ORDER_DETAIL[i]);
      ORDER_DETAIL_SMALL.push(must_col_content);
      must_col_content = null;
    }
    
    // console.log(must_col_content);
    //console.log(i);
  }
  //console.log(ORDER_DETAIL_SMALL);

  var buffer = xlsx.build([{name: "销售订单明细", data: ORDER_DETAIL_SMALL}]);
  fs.writeFileSync( "data/销售订单明细精简版.xlsx", buffer);

  vm.step = 160;
  return "";
};

// 第六步：补充数据到工作文件。
var fill_field_160 = function(){
  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse('data/销售订单明细精简版.xlsx');
  var order_info =  obj_sheet[0].data;
  MSG.put( " 销售订单明细精简版.XLSX  数据读入成功。");
  
  // 装入  [物料清单.XLSX]
  var obj_sheet2 = xlsx.parse('data/201511/物料清单.xlsx'); 
  var prod_info = obj_sheet2[0].data;
  MSG.put( " 物料清单.XLSX  数据读入成功。");
  // 物料清单中编码为 1001000202013110  16位   
  //   销售订单中  001001000202013110  18位
  // 需清洗数据。   

  var title_array = order_info[0];
  var index_prod_id = find_title_index(title_array, "物料号");
  var index_order_date = find_title_index(title_array, "创建日期");
  var index_cost = find_title_index(title_array, "成本单价");

  // 因为要跳过title，所以下标从 1 开始。
  for(var i=1; i<order_info.length; i++){
    var a_order = order_info[i];
    //console.log(i);
    
    var prod_id_in_order = a_order[index_prod_id];
    var date_in_order = a_order[index_order_date];
    //console.log(prod_id_in_order);
    //console.log(date_in_order);

    // 查「物料清单」表，取得成本价格
    var cost = getCost(prod_info, prod_id_in_order, date_in_order);
    //console.log(cost);
    // 成本 填入表格中。
    a_order[index_cost] = cost;

    //if( i>100 ) {break;}

  }

  var buffer = xlsx.build([{name: "销售订单明细(包含成本价)", data: order_info}]);
  fs.writeFileSync( "data/销售订单明细精简版(包含成本价).xlsx", buffer);

  vm.step = 170;
};

// 第七步：计算订单毛利。
var calc_gross_170 = function(){

  vm.step = 180;
};

// 第六步：加总物料毛利。
var calc_prod_180 = function(){

  vm.step = 190;
};

// 第六步：加总物料组毛利。
var calc_group_190 = function(){

  vm.step = 200;
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
};

var make_title = function( s_title ){
  var obj_title = {};
  obj_title.index = -1;
  obj_title.title = s_title;
  
  return obj_title;
};

var find_title_index = function( title_array, t_name){
  var a_index = _.indexOf( title_array, t_name );
  return a_index;
};

var getCost = function(prod_info, id, order_date){

  console.log(id);
    

  var a_index = -1;
  var prod_id = null;
  var id_a = id.trim();

  for(var i=2; i<prod_info.length; i++){

    prod_id = prod_info[i][0];
    if(_.isNumber(prod_id)){
      prod_id = ""+prod_id;
    }

    if( isblank(prod_id) ){
      a_index = -1
    }else{
      var id_b = "00" + prod_id.trim();   //  todo 临时解决编码长度不一致的问题。

      if( id_a == id_b ){
        a_index = i;
        
        break;
      }
    }
  }

  var cost = "";
  if( a_index < 0 ){
    console.log("Warning: index is -1.");
    cost = "";
  }else{
    cost = prod_info[i][8];
    console.log("-----------------" + cost);
  }
  //console.log(cost);
  return cost;
};

var find_prod = function(id){
  var prod = null;
  return prod;
};













