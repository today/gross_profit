/*!
 * index.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */

var fs = require('fs');
var xlsx = require("node-xlsx");

var ORDER_DETAIL = null;
var TAX_RATE = 1.17 ;


var init_100 = function(){

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
  
  // 程序运行所必须的库和配置文件
  var envlist = [];
  envlist.push("config.json");
  envlist.push("node_modules/node-xlsx/");
  envlist.push("node_modules/underscore/underscore-min.js");

  return true;
};

var check_status_120 = function(){
  
  var run_flag = true;

  // 程序运行所必须的数据源
  var srcFilelist = [];
  srcFilelist.push( vm.base_dir + "1月销售明细.XLSX");
  srcFilelist.push( vm.base_dir + "物料清单.XLSX");
  srcFilelist.push( vm.base_dir + "SCM客户明细.XLSX");
  //srcFilelist.push( vm.base_dir + "2015.11月计提并使用返利.XLSX");
  //srcFilelist.push( vm.base_dir + "2015.12促销品领用出库明细.XLSX");
  

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
    return true;
  }else{
    return false;
  }

  return "";
};

var check_src_130 = function(){
  return true;;
};

var load_order_detail_140 = function(){
  
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");

  var obj = null;
  
  obj = xlsx.parse( vm.base_dir + "1月销售明细.xlsx"); // 读入xlsx文件
  
  // 取出第一个sheet
  ORDER_DETAIL = obj[0].data; 

  // 清洗数据：去除空格
  trim_array_element(ORDER_DETAIL[0]); 
  console.log(ORDER_DETAIL[0]);

  MSG.put( "销售订单明细数据读入成功。");

  return true;;
};



var copy_order_detail_150 = function(){

  MSG.put( "数据较多，载入约需15秒。请耐心等待。");
  
  var must_col_title = [];
  must_col_title.push(make_title("实际交货数量"));
  must_col_title.push(make_title("物料描述"));
  must_col_title.push(make_title("交货完成状态"));
  must_col_title.push(make_title("销售价格"));
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
  must_col_title.push(make_title("内控成本价格"));
  //must_col_title.push(make_title("下游价返"));
  //must_col_title.push(make_title("促销费"));
  must_col_title.push(make_title("毛利"));
  must_col_title.push(make_title("毛利率"));

  
  var title_array = ORDER_DETAIL[0];
  // 把附加字段添加到 title 行中。
  title_array.push("内控成本价格");
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
  fs.writeFileSync( vm.base_dir + "销售订单明细精简版.xlsx", buffer);

  
  return true;;
};

// 第六步：补充数据到工作文件。
// 成本价格取数的时候  应该使用“内控成本价格” 
var fill_field_160 = function(){
  
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");
  

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse( vm.base_dir + "销售订单明细精简版.xlsx");
  var order_info =  obj_sheet[0].data;
  MSG.put( " 销售订单明细精简版.XLSX  数据读入成功。");
  
  // 装入  [物料清单.XLSX]
  var obj_sheet2 = xlsx.parse( vm.base_dir + "物料清单.xlsx"); 
  var prod_info = obj_sheet2[0].data;
  MSG.put( " 物料清单.XLSX  数据读入成功。");
  // 物料清单中编码为 1001000202013110  16位   
  //   销售订单中  001001000202013110  18位
  // 需清洗数据。   

  var title_array = order_info[0];
  var index_prod_id = find_title_index(title_array, "物料号");
  var index_order_date = find_title_index(title_array, "创建日期");
  var index_cost = find_title_index(title_array, "内控成本价格");

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
  fs.writeFileSync(  vm.base_dir + "销售订单明细精简版(包含成本价).xlsx", buffer);

  return true;
};

// 第七步：计算订单毛利。
/* 
毛利 = (销售价格-成本单价)*交货数量/1.17-下游价返-促销费

本步骤中，只计算  (销售价格-成本单价)*交货数量/1.17
*/
var calc_gross_170 = function(){
  
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse( vm.base_dir + "销售订单明细精简版(包含成本价).xlsx");
  var order_info =  obj_sheet[0].data;
  MSG.put( " 销售订单明细精简版(包含成本价).xlsx  数据读入成功。");

  var title_array = order_info[0];
  var index_price = find_title_index(title_array, "销售价格");
  var index_delivery_count = find_title_index(title_array, "实际交货数量");
  var index_cost = find_title_index(title_array, "内控成本价格");
  var index_gross = find_title_index(title_array, "毛利");
  var index_gross_rate = find_title_index(title_array, "毛利率");


  // 因为要跳过title，所以下标从 1 开始。
  for(var i=1; i<order_info.length; i++){
    var a_order = order_info[i];

    var price = a_order[index_price];
    var delivery_count = a_order[index_delivery_count];
    var cost = a_order[index_cost];

    if( -1 === cost ){
      console.log("数据错，忽略。");
    }
    else if( 0 < cost ){
      a_order[index_gross] = (price - cost) * delivery_count / TAX_RATE ;
      a_order[index_gross_rate] = a_order[index_gross] / price * 100;
    }else{
      ERR_MSG.put("数据出错：成本价数据异常。行数：" + (i+1) + " 成本价：" + cost );
    }

  }

  var buffer = xlsx.build([{name: "销售毛利", data: order_info}]);
  fs.writeFileSync(  vm.base_dir + "销售毛利.xlsx", buffer);

  return true;;
};

// 第六步：加总物料毛利。
var calc_prod_180 = function(){
  
  MSG.put( "数据较多，载入约需15秒。请耐心等待。");

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse( vm.base_dir + "销售毛利.xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 销售毛利.xlsx  数据读入成功。");

  // 获取title
  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("单物料毛利");

  var index_count = find_title_index(title_array, "实际交货数量");
  var index_prod_id = find_title_index(title_array, "物料号");
  var index_prod_group_id = find_title_index(title_array, "物料组");
  var index_gross = find_title_index(title_array, "毛利");
  var index_gross_sum = find_title_index(title_array, "单物料毛利");
  
  // 这将是一个命名数组，也就是类似java中的hashArray，或者Py中的dict
  var gross_sum = [];
  // 加总毛利
  for(var i=1; i<gross_info.length; i++){
    var a_order = gross_info[i];

    var count = a_order[index_count];
    var prod_id = a_order[index_prod_id];
    var gross = a_order[index_gross];

    
    if( gross_sum[prod_id] ){
      // 单品汇总信息已经存在
      a_temp = gross_sum[prod_id];
      // 累加毛利
      a_temp[index_gross_sum] += gross;
      // 累加数量
      a_temp[index_count] += count;
    }
    else{
      // 单品汇总信息 尚未存在
      a_order.push(gross);
      gross_sum[prod_id] = a_order;
    }
  }

  var prod_gross = [];
  // 填充标题栏
  prod_gross.push(title_array);
  // 填充数据
  for(var key in gross_sum){
      prod_gross.push(gross_sum[key]);
  }

  // 标识不需要的列
  var index_will_delete = [];
  index_will_delete.push( find_title_index(title_array, "销售数量") );
  index_will_delete.push( find_title_index(title_array, "订单类型") );
  index_will_delete.push( find_title_index(title_array, "交货完成状态") );
  index_will_delete.push( find_title_index(title_array, "实际交货日期") );
  index_will_delete.push( find_title_index(title_array, "渠道") );
  index_will_delete.push( find_title_index(title_array, "客户") );
  index_will_delete.push( find_title_index(title_array, "客户名称") );
  index_will_delete.push( find_title_index(title_array, "城市") );
  index_will_delete.push( find_title_index(title_array, "客户参考号") );
  index_will_delete.push( find_title_index(title_array, "创建日期") );
  index_will_delete.push( find_title_index(title_array, "毛利") );
  index_will_delete.push( find_title_index(title_array, "毛利率") );
  // 清除不需要的列
  var thin_gross = del_col_from_array(prod_gross, index_will_delete);


  var buffer = xlsx.build([{name: "物料毛利", data: thin_gross}]);
  fs.writeFileSync(  vm.base_dir + "物料毛利.xlsx", buffer);

  return true;
};

// 第六步：加总物料组毛利。
var calc_group_190 = function(){
  
  MSG.put( "数据较多，载入约需5秒。请耐心等待。");

  // 装入  [销售订单明细精简版.xlsx]
  var obj_sheet = xlsx.parse( vm.base_dir + "物料毛利.xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 物料毛利.xlsx  数据读入成功。");

  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("物料组毛利");

  var index_count = find_title_index(title_array, "实际交货数量");
  var index_prod_group_id = find_title_index(title_array, "物料组");
  var index_gross = find_title_index(title_array, "单物料毛利");
  var index_group_gross = find_title_index(title_array, "物料组毛利");

  var group_sum = [];
  // 加总物料组毛利
  for(var i=1; i<gross_info.length; i++){
    var a_prod_gross = gross_info[i];

    var count = a_prod_gross[index_count];
    var group_id = a_prod_gross[index_prod_group_id];
    var gross = a_prod_gross[index_gross];

    if( group_sum[group_id] ){
      // 单品汇总信息已经存在
      a_temp = group_sum[group_id];
      // 累加毛利
      a_temp[index_group_gross] += gross;
      // 累加数量
      a_temp[index_count] += count;
    }
    else{
      // 单品汇总信息 尚未存在
      a_prod_gross.push(gross);
      group_sum[group_id] = a_prod_gross;
    }
  }

  var group_gross_sum = [];
  // 填充标题栏
  group_gross_sum.push(title_array);
  // 填充数据
  for(var key in group_sum){
      group_gross_sum.push(group_sum[key]);
  } 

  // 标识不需要的列
  var index_will_delete = [];
  index_will_delete.push( find_title_index(title_array, "物料描述") );
  index_will_delete.push( find_title_index(title_array, "销售价格") );
  index_will_delete.push( find_title_index(title_array, "物料号") );
  index_will_delete.push( find_title_index(title_array, "成本单价") );
  index_will_delete.push( find_title_index(title_array, "单物料毛利") );
  // 清除不需要的列
  var thin_gross = del_col_from_array(group_gross_sum, index_will_delete);

  var buffer = xlsx.build([{name: "物料组毛利", data: thin_gross}]);
  fs.writeFileSync(  vm.base_dir + "物料组毛利.xlsx", buffer);

  return true;
};


var calc_branch_200 = function(){
  
  return true;
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

  // TOOD 要处理成根据title找出index的模式。
  //var index_cost = find_title_index(prod_info[0], "内控成本价格");

  var cost = "";
  if( a_index < 0 ){
    console.log("Warning: index is -1.");
    cost = "";
  }else{
    cost = prod_info[i][10];
    //console.log("-----------------" + cost);
  }
  //console.log(cost);
  return cost;
};

var find_prod = function(id){
  var prod = null;
  return prod;
};















