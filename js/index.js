/*!
 * index.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */

var fs = require('fs');
var _ = require('underscore');
var xlsx = require("node-xlsx");

var ORDER_DETAIL = null;
var TAX_RATE = 1.17 ;


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
  vm.step = 141;

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



var copy_order_detail_150 = function(){
  console.log("copy_order_detail_150");
  vm.step = 151;
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");
  
  
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
  must_col_title.push(make_title("毛利"));
  must_col_title.push(make_title("毛利率"));

  
  var title_array = ORDER_DETAIL[0];
  // 把附加字段添加到 title 行中。
  title_array.push("成本单价");
  //title_array.push("下游价返");
  //title_array.push("促销费");
  title_array.push("毛利");
  title_array.push("毛利率");

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

  // 进行数据清洗，物料号，把18位的编码缩减到16位。
  var prod_id_index = find_title_index(must_col_title, "物料号");
  for(var i=1;i<ORDER_DETAIL_SMALL.length; i++){
    var prod_id_temp = ORDER_DETAIL_SMALL[i][prod_id_index];
    if( 18 === prod_id_temp.length ){
      ORDER_DETAIL_SMALL[i][prod_id_index] = prod_id_temp.substring(2);
    }else{
      ERR_MSG.put("数据出错：订单表中的物料号长度不是20。行数：" + i + " 物料号：" + prod_id_temp );
    }
  }


  //console.log(ORDER_DETAIL_SMALL);

  var buffer = xlsx.build([{name: "销售订单明细", data: ORDER_DETAIL_SMALL}]);
  fs.writeFileSync( "data/销售订单明细精简版.xlsx", buffer);

  vm.step = 160;
  return "";
};

// 第六步：补充数据到工作文件。
var fill_field_160 = function(){
  console.log("fill_field_160");
  vm.step = 161;
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");
  

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
    
    if( undefined === cost ){
      ERR_MSG.put("数据出错：物料表中的成本价格未填写。行数：" + (i+1) + " 物料号：" + prod_id_in_order );
      a_order[index_cost] = -1;
    }else{
      // 成本 填入表格中。
      a_order[index_cost] = cost;
      //console.log(cost);
    }

  }

  var buffer = xlsx.build([{name: "销售订单明细(包含成本价)", data: order_info}]);
  fs.writeFileSync( "data/销售订单明细精简版(包含成本价).xlsx", buffer);

  vm.step = 170;
};

// 第七步：计算订单毛利。
/* 
毛利 = （销售收入-成本单价*交货数量）/1.17-下游价返-促销费

本步骤中，只计算  （销售收入-成本单价*交货数量）/1.17
*/
var calc_gross_170 = function(){
  console.log("calc_gross_170");
  vm.step = 171;
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse('data/销售订单明细精简版(包含成本价).xlsx');
  var order_info =  obj_sheet[0].data;
  MSG.put( " 销售订单明细精简版(包含成本价).xlsx  数据读入成功。");

  var title_array = order_info[0];
  var index_total = find_title_index(title_array, "销售金额");
  var index_delivery_count = find_title_index(title_array, "实际交货数量");
  var index_cost = find_title_index(title_array, "成本单价");
  var index_gross = find_title_index(title_array, "毛利");
  var index_gross_rate = find_title_index(title_array, "毛利率");


  // 因为要跳过title，所以下标从 1 开始。
  for(var i=1; i<order_info.length; i++){
    var a_order = order_info[i];

    var total = a_order[index_total];
    var delivery_count = a_order[index_delivery_count];
    var cost = a_order[index_cost];

    if( -1 === cost ){
      console.log("数据错，忽略。");
    }
    else if( 0 < cost ){
      a_order[index_gross] = ( total - (cost * delivery_count) ) / TAX_RATE ;
      a_order[index_gross_rate] = a_order[index_gross] / total * 100;
    }else{
      ERR_MSG.put("数据出错：成本价数据异常。行数：" + (i+1) + " 成本价：" + cost );
    }

  }

  var buffer = xlsx.build([{name: "销售订单明细(包含成本价)", data: order_info}]);
  fs.writeFileSync( "data/销售订单明细精简版(包含成本价).xlsx", buffer);

  vm.step = 180;
};

// 第六步：加总物料毛利。
var calc_prod_180 = function(){
  console.log("calc_prod_180");
  vm.step = 181;
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse('data/销售订单明细精简版(包含成本价).xlsx');
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 销售订单明细精简版(包含成本价).xlsx  数据读入成功。");


  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("单物料毛利");

  var index_prod_id = find_title_index(title_array, "物料号");
  var index_prod_group_id = find_title_index(title_array, "物料组");
  //var index_total = find_title_index(title_array, "销售金额");
  var index_gross = find_title_index(title_array, "毛利");
  var index_gross_sum = find_title_index(title_array, "单物料毛利");
  
  

  var gross_sum = [];
  // 加总毛利
  for(var i=1; i<gross_info.length; i++){
    var a_order = gross_info[i];

    var prod_id = a_order[index_prod_id];
    var gross = a_order[index_gross];

    if( gross_sum[prod_id] ){
      // 单品汇总信息已经存在
      a_temp = gross_sum[prod_id];
      // 累加毛利
      a_temp[index_gross_sum] += gross;
    }
    else{
      // 单品汇总信息 尚未存在
      a_order.push(gross);
      gross_sum[prod_id] = a_order;
    }

  }

  var prod_gross = [];
  // 填充标题栏数据
  prod_gross.push(title_array);
  for(var key in gross_sum){
      prod_gross.push(gross_sum[key]);
  } 


  var buffer = xlsx.build([{name: "物料毛利", data: prod_gross}]);
  fs.writeFileSync( "data/物料毛利.xlsx", buffer);

  vm.step = 190;
};

// 第六步：加总物料组毛利。
var calc_group_190 = function(){
  console.log("calc_group_190");
  vm.step = 191;
  MSG.put( "数据较多，载入约需5秒。请耐心等待。");

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse('data/物料毛利.xlsx');
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 物料毛利.xlsx  数据读入成功。");

  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("物料组毛利");

  var index_prod_group_id = find_title_index(title_array, "物料组");
  var index_gross = find_title_index(title_array, "单物料毛利");
  var index_group_gross = find_title_index(title_array, "物料组毛利");

  var group_sum = [];
  // 加总物料组毛利
  for(var i=1; i<gross_info.length; i++){
    var a_prod_gross = gross_info[i];

    var group_id = a_prod_gross[index_prod_group_id];
    var gross = a_prod_gross[index_gross];

    if( group_sum[group_id] ){
      // 单品汇总信息已经存在
      a_temp = group_sum[group_id];
      // 累加毛利
      a_temp[index_group_gross] += gross;
    }
    else{
      // 单品汇总信息 尚未存在
      a_prod_gross.push(gross);
      group_sum[group_id] = a_prod_gross;
    }

  }

  var group_gross_sum = [];
  // 填充标题栏数据
  group_gross_sum.push(title_array);
  for(var key in group_sum){

      group_gross_sum.push(group_sum[key]);
  } 

  var buffer = xlsx.build([{name: "物料组毛利", data: group_gross_sum}]);
  fs.writeFileSync( "data/物料组毛利.xlsx", buffer);

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
  //console.log(id);
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
      var id_b = prod_id.trim(); 

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
    //console.log("-----------------" + cost);
  }
  //console.log(cost);
  return cost;
};

var find_prod = function(id){
  var prod = null;
  return prod;
};













