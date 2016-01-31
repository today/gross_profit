/*!
 * index.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */
var xlsx = require("node-xlsx");

var ORDER_DETAIL = null;
var TAX_RATE = 1.17 ;

var init_100 = function(){

  MSG.put("系统启动。");

}

var check_env_110 = function(){
  
  // 程序运行所必须的库和配置文件
  var envlist = [];
  envlist.push("config.json");
  envlist.push("node_modules/node-xlsx/");
  envlist.push("node_modules/underscore/underscore-min.js");

  return true;
};

var select_file_120 = function(){
  
  if(fs.existsSync(vm.sales_filename)) {
    console.log('销售记录文件存在');
    
  } else {
    console.log('销售记录文件不存在');
    document.getElementById("file_src").click();
  }
  return true;
};

var check_src_130 = function(){
  var run_flag = true;

  // 设置base_dir
  var temp_path = document.getElementById("file_src").value;
  vm.base_dir = path.dirname(temp_path ) + "/";
  console.log("base_dir: " + vm.base_dir );

  // 充当数据源的文件 的文件名 的关键字
  vm.src_files_flag.push("销售订单明细");
  vm.src_files_flag.push("物料清单");
  vm.src_files_flag.push("SCM客户明细");
  vm.src_files_flag.push("内控成本价格变动");
  //vm.src_files_flag.push("xxxxxxx");

  // 程序运行所必须的数据源
  vm.src_files = find_src_file(vm.base_dir, vm.src_files_flag);
  //srcFilelist.push( vm.base_dir + "2015.11月计提并使用返利.XLSX");
  //srcFilelist.push( vm.base_dir + "2015.12促销品领用出库明细.XLSX");
  
  console.log( vm.src_files);

  for( temp_name in vm.src_files){
    if( undefined === vm.src_files[temp_name]){
      ERR_MSG.put( "输入文件不全。缺少：" + temp_name );
      run_flag = false;
    }
  }


  return run_flag;;
};

var load_order_detail_140 = function(){

  var obj = null;
  
  obj = xlsx.parse( vm.base_dir + vm.src_files['销售订单明细'] ); // 读入xlsx文件
  
  // 取出第一个sheet
  ORDER_DETAIL = obj[0].data; 

  // 清洗数据：去除空格
  trim_array_element(ORDER_DETAIL[0]); 
  console.log(ORDER_DETAIL[0]);

  MSG.put( "销售订单明细   数据读入成功。");

  return true;;
};



var copy_order_detail_150 = function(){
  
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
  must_col_title.push(make_title("物料组描述"));
  must_col_title.push(make_title("库存地点"));
  must_col_title.push(make_title("实际交货日期"));
  must_col_title.push(make_title("城市"));
  must_col_title.push(make_title("客户参考号"));
  must_col_title.push(make_title("创建日期"));
  // 以下数据，要从其他文件中获得
  must_col_title.push(make_title("内控成本价格"));
  //must_col_title.push(make_title("下游价返"));
  //must_col_title.push(make_title("促销费"));
  must_col_title.push(make_title("销售收入"));
  must_col_title.push(make_title("毛利"));
  must_col_title.push(make_title("毛利率"));

  
  var title_array = ORDER_DETAIL[0];
  // 把附加字段添加到 title 行中。
  title_array.push("内控成本价格");
  title_array.push("销售收入");
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

  // 进行数据清洗，物料号，把18位的编码缩减到16位。客户编码，删除前面的两个零。
  var prod_id_index = find_title_index(must_col_title, "物料号");
  var custom_id_index = find_title_index(must_col_title, "客户");
  for(var i=1;i<ORDER_DETAIL_SMALL.length; i++){
    var prod_id_temp = ORDER_DETAIL_SMALL[i][prod_id_index];
    var custom_id_temp = ORDER_DETAIL_SMALL[i][custom_id_index];

    if( 18 === prod_id_temp.length ){
      ORDER_DETAIL_SMALL[i][prod_id_index] = prod_id_temp.substring(2);
      ORDER_DETAIL_SMALL[i][custom_id_index] = custom_id_temp.substring(2);
    }else{
      ERR_MSG.put("数据出错：订单表中的物料号长度不是20。行数：" + i + " 物料号：" + prod_id_temp );
    }

    ORDER_DETAIL_SMALL[i][custom_id_index] = custom_id_temp.substring(2);
    
  }

  var buffer = xlsx.build([{name: "销售订单明细", data: ORDER_DETAIL_SMALL}]);
  fs.writeFileSync( vm.base_dir + "中间文件_销售数据.xlsx", buffer);

  
  return true;;
};

// 第六步：补充数据到工作文件。
// 成本价格取数的时候  应该使用“内控成本价格” 
var fill_field_160 = function(){
  
  // 装入  [销售订单明细]
  var obj_sheet = xlsx.parse( vm.base_dir + "中间文件_销售数据.xlsx");
  var order_info =  obj_sheet[0].data;
  MSG.put( " 中间文件_销售数据.XLSX  数据读入成功。");
  
  // 获得物料数据。
  var prod_info = getProd_info();


  // 销售订单表
  var title_array_3 = order_info[0];
  console.log(title_array_3);
  var index_prod_id = find_title_index(title_array_3, "物料号");
  var index_order_date = find_title_index(title_array_3, "创建日期");
  var index_cost = find_title_index(title_array_3, "内控成本价格");

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
      ERR_MSG.put("数据出错：物料表中的成本价格未填写。行数：\t" + (i+1) + "\t物料号：\t" + prod_id_in_order );
      a_order[index_cost] = -1;
    }else{
      // 成本 填入表格中。
      a_order[index_cost] = cost;
      //console.log(cost);
    }

  }

  var buffer = xlsx.build([{name: "中间文件_销售数据(包含成本价)", data: order_info}]);
  fs.writeFileSync(  vm.base_dir + "中间文件_销售数据(包含成本价).xlsx", buffer);

  return true;
};

// 第七步：计算订单毛利。
/* 
毛利 = (销售价格-成本单价)*交货数量/1.17-下游价返-促销费

本步骤中，只计算  (销售价格-成本单价)*交货数量/1.17
*/
var calc_gross_170 = function(){

  var obj_sheet = xlsx.parse( vm.base_dir + "中间文件_销售数据(包含成本价).xlsx");
  var order_info =  obj_sheet[0].data;
  MSG.put( " 中间文件_销售数据(包含成本价).xlsx  数据读入成功。");

  var title_array = order_info[0];
  var index_price = find_title_index(title_array, "销售价格");
  var index_delivery_count = find_title_index(title_array, "实际交货数量");
  var index_cost = find_title_index(title_array, "内控成本价格");
  var index_income = find_title_index(title_array, "销售收入");
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
      var cost_sum = cost * delivery_count;
      var income_sum = price * delivery_count
      a_order[index_income] = income_sum;
      a_order[index_gross] = (income_sum - cost_sum) / TAX_RATE ;
      a_order[index_gross_rate] = a_order[index_gross] / cost_sum * 100;
    }else{
      ERR_MSG.put("数据出错：成本价数据异常。行数：" + (i+1) + " 成本价：" + cost );
    }

  }

  var buffer = xlsx.build([{name: "计算结果_销售毛利", data: order_info}]);
  fs.writeFileSync(  vm.base_dir + "计算结果_销售毛利.xlsx", buffer);

  return true;;
};

// 第六步：加总物料毛利。
var calc_prod_180 = function(){
  
  var obj_sheet = xlsx.parse( vm.base_dir + "计算结果_销售毛利.xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 计算结果_销售毛利.xlsx  数据读入成功。");

  // 获取title
  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("单物料毛利");

  var index_count = find_title_index(title_array, "实际交货数量");
  var index_prod_id = find_title_index(title_array, "物料号");
  var index_prod_group_id = find_title_index(title_array, "物料组");
  var index_income = find_title_index(title_array, "销售收入");
  var index_gross = find_title_index(title_array, "毛利");
  var index_gross_sum = find_title_index(title_array, "单物料毛利");
  
  // 这将是一个命名数组，也就是类似java中的hashArray，或者Py中的dict
  var gross_sum = [];
  // 加总毛利
  for(var i=1; i<gross_info.length; i++){
    var a_order = gross_info[i];

    var prod_id = a_order[index_prod_id];
    var count = a_order[index_count];
    var income = a_order[index_income];
    var gross = a_order[index_gross];
    
    if( gross_sum[prod_id] ){
      // 单品汇总信息已经存在
      a_temp = gross_sum[prod_id];
      // 累加数量
      a_temp[index_count] += count;
      // 累加收入
      a_temp[index_income] += income;
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


  var buffer = xlsx.build([{name: "计算结果_物料毛利", data: thin_gross}]);
  fs.writeFileSync(  vm.base_dir + "计算结果_物料毛利.xlsx", buffer);

  return true;
};

// 第六步：加总物料组毛利。
var calc_group_190 = function(){
  
  MSG.put( "数据较多，载入约需5秒。请耐心等待。");

  var obj_sheet = xlsx.parse( vm.base_dir + "计算结果_物料毛利.xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 计算结果_物料毛利.xlsx  数据读入成功。");

  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("物料组毛利");

  var index_count = find_title_index(title_array, "实际交货数量");
  var index_prod_group_id = find_title_index(title_array, "物料组");
  var index_gross = find_title_index(title_array, "单物料毛利");
  var index_income = find_title_index(title_array, "销售收入");
  var index_group_gross = find_title_index(title_array, "物料组毛利");

  var group_sum = [];
  // 加总物料组毛利
  for(var i=1; i<gross_info.length; i++){
    var a_prod_gross = gross_info[i];

    var count = a_prod_gross[index_count];
    var group_id = a_prod_gross[index_prod_group_id];
    var income = a_prod_gross[index_income];
    var gross = a_prod_gross[index_gross];

    if( group_sum[group_id] ){
      // 单品汇总信息已经存在
      a_temp = group_sum[group_id];
      // 累加数量
      a_temp[index_count] += count;
      // 累加收入
      a_temp[index_income] += income;
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

  var buffer = xlsx.build([{name: "计算结果_物料组毛利", data: thin_gross}]);
  fs.writeFileSync(  vm.base_dir + "计算结果_物料组毛利.xlsx", buffer);

  return true;
};



var getProd_info = function(){
  // 装入  [物料清单.XLSX]
  // 物料清单中编码为 1001000202013110  16位   
  //   销售订单中  001001000202013110  18位
  // 需清洗数据。
  var obj_sheet2 = xlsx.parse( vm.base_dir + vm.src_files['物料清单']); 
  var prod_info = obj_sheet2[0].data;
  MSG.put( " 物料清单  数据读入成功。");

  for(var i=1; i<prod_info.length; i++){
    prod_info[i].push(40000);
  }

  // 取出必要的列
  var title_array = prod_info[0];
  title_array.push("开始变动日期")

  console.log(title_array);
  var index_must = [];
  index_must.push( find_title_index(title_array, "物料号") );
  index_must.push( find_title_index(title_array, "物料描述") );
  index_must.push( find_title_index(title_array, "内控成本价格") );
  index_must.push( find_title_index(title_array, "开始变动日期") );
  var prod_must_col = select_col_from_array(prod_info, index_must);
  
  // 装入   内控成本价格变动--1月汇总.xlsx
  var obj_sheet3 = xlsx.parse( vm.base_dir + vm.src_files['内控成本价格变动']); 
  var price_history = obj_sheet3[0].data;
  MSG.put( " 内控成本价格变动  数据读入成功。");
  // 取出必要的列
  var title_array_2 = price_history[0];

  var index_must_2 = [];
  index_must_2.push( find_title_index(title_array_2, "物料编码") );
  index_must_2.push( find_title_index(title_array_2, "产品描述") );
  index_must_2.push( find_title_index(title_array_2, "内控成本价格") );
  index_must_2.push( find_title_index(title_array_2, "开始变动日期") );
  var cost_history_col = select_col_from_array(price_history, index_must_2);

  // 合并价格变动数组到物料数组中
  for(var i=1; i<cost_history_col.length; i++){
    prod_must_col.push(cost_history_col[i]);
  }

  console.log(index_must);
  console.log(index_must_2);
  return prod_must_col;
}

var fill_branch_200 = function(){

  var obj_sheet = xlsx.parse( vm.base_dir + "计算结果_销售毛利.xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 计算结果_销售毛利.xlsx  数据读入成功。");

  // 取出必要的列
  var title_array = gross_info[0];
  title_array.push("渠道归属")
  console.log(title_array);

  var index_custom_id = find_title_index(title_array, "客户");
  var index_warehouse_id = find_title_index(title_array, "库存地点");
  var index_branch = find_title_index(title_array, "渠道归属");
  
  // 第一步，筛选出仓库（表头为：库存地点）是 P012 的，全部是电子渠道。
  // 第二步，仓库是 P001~P011 ，并且客户编码(表头为：客户)是 001 开头，是自有渠道。
  // 第三步，仓库是 1010  开头，并且客户编码 002 开头，是零售渠道。
  // 第四步，仓库是 任意 , 客户编码是 8 开头，是分销渠道。
  // 第五步，如果仍然有剩余数据，报错。
  // 注：客户编码也需要做数据清洗，已经在前面做了。
  var temp_array = [];
  for(var i=1; i<gross_info.length; i++){
    var temp_order = gross_info[i];
    var temp_id = temp_order[index_warehouse_id];
    if( 'P012' === temp_id.substring(0,4) ){
      temp_order[index_branch] = '电子渠道';
    }else{
      temp_array.push(temp_order);
    }
  }

  var temp_array2 = [];
  for(var i=1; i<temp_array.length; i++){
    var temp_order = temp_array[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    if( 'P0' === temp_id.substring(0,2) && '001'===custom_id.substring(0,3) ){
      temp_order[index_branch] = '自有渠道';
    }else{
      temp_array2.push(temp_order);
    }
  }

  var temp_array3 = [];
  for(var i=1; i<temp_array2.length; i++){
    var temp_order = temp_array2[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    if( '1010' === temp_id.substring(0,4) && '002'===custom_id.substring(0,3) ){
      temp_order[index_branch] = '零售渠道';
    }else{
      temp_array3.push(temp_order);
    }
  }

  var temp_array4 = [];
  for(var i=1; i<temp_array3.length; i++){
    var temp_order = temp_array3[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    if( '8'===custom_id.substring(0,1) ){
      temp_order[index_branch] = '分销渠道';
    }else{
      temp_array4.push(temp_order);
    }
  }

  if( temp_array4.length > 0 ){
    console.log("数据出错：发现无渠道归属的订单。");
    //console.log(temp_array4);
  }

  var buffer = xlsx.build([{name: "销售数据(渠道归属)", data: gross_info}]);
  fs.writeFileSync(  vm.base_dir + "中间文件_销售数据(渠道归属).xlsx", buffer);


  return true;
};

var fill_city_210 = function(){
  var obj_sheet = xlsx.parse( vm.base_dir + "中间文件_销售数据(渠道归属).xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 中间文件_销售数据(渠道归属).xlsx  数据读入成功。");

  // 取出必要的列
  var title_array = gross_info[0];
  title_array.push("地市归属");
  console.log(title_array);
  var index_custom_id = find_title_index(title_array, "客户");
  var index_city = find_title_index(title_array, "地市归属");
  

  var obj_sheet2 = xlsx.parse( vm.base_dir + vm.src_files['SCM客户明细']);
  var custom_info =  obj_sheet2[0].data;
  MSG.put( " SCM客户明细  数据读入成功。");

  for(var i=1; i<gross_info.length; i++ ){
    var order = gross_info[i];
    var city = getCity(custom_info, order[index_custom_id]);
    //console.log(city);
    order[index_city] = city;

    //if(i>10)break;
  }

  var buffer = xlsx.build([{name: "销售数据(渠道归属和地市归属)", data: gross_info}]);
  fs.writeFileSync(  vm.base_dir + "中间文件_销售数据(渠道归属和地市归属).xlsx", buffer);
  
  return true;
};

var calc_branch_city_220 = function(){
  var obj_sheet = xlsx.parse( vm.base_dir + "中间文件_销售数据(渠道归属和地市归属).xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 中间文件_销售数据(渠道归属和地市归属).xlsx  数据读入成功。");

  // 取出必要的列
  var title_array = gross_info[0];
  console.log(title_array);
  
  var index_count = find_title_index(title_array, "实际交货数量");
  var index_income = find_title_index(title_array, "销售收入");
  var index_gross = find_title_index(title_array, "毛利");
  var index_branch = find_title_index(title_array, "渠道归属");
  var index_city = find_title_index(title_array, "地市归属");
  
  // 结果文件的列 title
  var title_array_dest = [];
  title_array_dest.push("");
  title_array_dest.push("自有渠道销量");
  title_array_dest.push("自有渠道收入");
  title_array_dest.push("自有渠道毛利");  
  title_array_dest.push("分销渠道销量");  
  title_array_dest.push("分销渠道收入");  
  title_array_dest.push("分销渠道毛利");  
  title_array_dest.push("电子渠道销量");  
  title_array_dest.push("电子渠道收入");  
  title_array_dest.push("电子渠道毛利");  
  title_array_dest.push("零售渠道销量");  
  title_array_dest.push("零售渠道收入");  
  title_array_dest.push("零售渠道毛利");  
  title_array_dest.push("合计销量");  
  title_array_dest.push("合计收入");  
  title_array_dest.push("合计毛利");

  var make_summary_line = function(title){ 
    var data = ["",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
    data[0] = title;
    return data;
  };
  //结果数据
  var data_array = [];
  data_array["西安城区"] = make_summary_line('西安城区');
  data_array["西安郊县"] = make_summary_line('西安郊县');
  data_array["咸阳"]    = make_summary_line('咸阳');
  data_array["宝鸡"]    = make_summary_line('宝鸡');
  data_array["渭南"]    = make_summary_line('渭南');
  data_array["铜川"]    = make_summary_line('铜川');
  data_array["延安"]    = make_summary_line('延安');
  data_array["榆林"]    = make_summary_line('榆林');
  data_array["汉中"]    = make_summary_line('汉中');
  data_array["安康"]    = make_summary_line('安康');
  data_array["商洛"]    = make_summary_line('商洛');
  data_array["其他"]    = make_summary_line('其他');
  data_array["零售中心"] = make_summary_line('零售中心');
  data_array["数据错误"]    = make_summary_line('数据错误');

  // 加总数量，收入，毛利  的函数。
  var calc_summary = function( order, data){
    var count = order[index_count];
    var income = order[index_income];
    var gross = order[index_gross];

    if( count === undefined ){
      count = 0;
    }  
    if( income === undefined ){
      income = 0;
    }
    if( gross === undefined ){
      gross = 0;
    }

    var branch = order[index_branch]

    var offset = 0;
    if( "自有渠道" === branch ){
      offset = 0;
    }else if ( "分销渠道" === branch ){
      offset = 3;
    }else if ( "电子渠道" === branch ){
      offset = 6;
    }else if ( "零售渠道" === branch ){
      offset = 9;
    }

    data[offset+1] += count;
    data[offset+2] += income;
    data[offset+3] += gross;

    data[13] += count;
    data[14] += income;
    data[15] += gross;

  };

  for(var i=1; i<gross_info.length; i++ ){
    var order = gross_info[i];
    var city = order[index_city];
    //console.log(city); 
    if( data_array[city] ){
      calc_summary(order, data_array[city]);
    }else{
      calc_summary(order, data_array['数据错误']);
    }
  }

  var full_data = [];
  full_data.push(title_array_dest);
  full_data.push(data_array["西安城区"] );
  full_data.push(data_array["西安郊县"] );
  full_data.push(data_array["咸阳"]    );
  full_data.push(data_array["宝鸡"]    );
  full_data.push(data_array["渭南"]    );
  full_data.push(data_array["铜川"]    );
  full_data.push(data_array["延安"]    );
  full_data.push(data_array["榆林"]    );
  full_data.push(data_array["汉中"]    );
  full_data.push(data_array["安康"]    );
  full_data.push(data_array["商洛"]    );
  full_data.push(data_array["其他"]    );
  full_data.push(data_array["数据错误"]    );
  //full_data.push(data_array["零售中心"] );
  
  var buffer = xlsx.build([{name: "地市渠道销量收入毛利汇总表", data: full_data}]);
  fs.writeFileSync(  vm.base_dir + "计算结果_地市渠道销量收入毛利汇总表.xlsx", buffer);
  return true;
};

var success_230 = function () {
  return true;
}

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

var getCity = function(custom_info, custom_id){

  var title_array = custom_info[0];
  //console.log(title_array);
  custom_no_index = find_title_index(title_array, "客户");
  city_index = find_title_index(title_array, "地市");
  //console.log(custom_no_index);
  //console.log(city_index);
  var ret_city = null;

  for(var i=1; i<custom_info.length; i++){
    var custom = custom_info[i];
    var no = custom[custom_no_index];
    var city = custom[city_index];

    //console.log(custom_id);
    //console.log(no);
    if( custom_id == no ){
      if( city.indexOf("西安城区") > -1 ){
        ret_city = "西安城区";
      }else if( city.indexOf("西安郊县") > -1 ){
        ret_city = "西安郊县";
      }else if( city.indexOf("咸阳") > -1 ){
        ret_city = "咸阳";
      }else if( city.indexOf("宝鸡") > -1 ){
        ret_city = "宝鸡";
      }else if( city.indexOf("渭南") > -1 ){
        ret_city = "渭南";
      }else if( city.indexOf("铜川") > -1 ){
        ret_city = "铜川";
      }else if( city.indexOf("延安") > -1 ){
        ret_city = "延安";
      }else if( city.indexOf("榆林") > -1 ){
        ret_city = "榆林";
      }else if( city.indexOf("汉中") > -1 ){
        ret_city = "汉中";
      }else if( city.indexOf("安康") > -1 ){
        ret_city = "安康";
      }else if( city.indexOf("商洛") > -1 ){
        ret_city = "商洛";
      }else if( city.indexOf("其他") > -1 ){
        ret_city = "其他";
      }else if( city.indexOf("零售中心") > -1 ){
        ret_city = "零售中心";
      }else{
        ret_city = "出错";
      }
      break;
    }
  }
  return ret_city;
}

var getCost = function(prod_info, id, order_date){
  //console.log(id);
  var temp_array = [];
  var a_index = -1;
  var prod_id = null;
  var id_a = id.trim();

  var title_array = prod_info[0];
  id_index = find_title_index(title_array, "物料号");
  cost_index = find_title_index(title_array, "内控成本价格");
  date_index = find_title_index(title_array, "开始变动日期");

  for(var i=2; i<prod_info.length; i++){

    prod_id = prod_info[i][id_index];
    if(_.isNumber(prod_id)){
      prod_id = ""+prod_id;
    }

    if( isblank(prod_id) ){
      a_index = -1;
      console.log(id);
    }else{
      var id_b = prod_id.trim(); 

      if( id_a == id_b ){
        temp_array.push(prod_info[i]);
      }
    }
  }

  // 看看哪个价钱是当前的价钱
  var cost = -1;
  if( temp_array.length === 0 ){
    console.log("Warning: cost not found.");
    cost = "";
  }
  else if( temp_array.length === 1 ){
    cost = temp_array[0][cost_index];
  }
  else{
    // 先排序。
    //console.log("出现多个成本价，找日期最接近的那个。 order_date= " + order_date );
    //console.log(temp_array);
    temp_array = _.sortBy(temp_array, function(num){ return num[date_index]; });
    //console.log(temp_array);

    var temp_date = 0;
    for(var i=1;i<temp_array.length; i++){
      var c1 = temp_array[i-1];
      var c2 = temp_array[i];
      if( order_date >= c1[date_index] && order_date < c2[date_index] ){
        cost = temp_array[i-1][cost_index];
        // console.log("------------");
        // console.log(c1[date_index]);
        // console.log(order_date);
        // console.log(c2[date_index]);
        // console.log("------------");
        break;
      }
      else if( i === temp_array.length-1 && order_date > c2[date_index]){
        cost = temp_array[i][cost_index];
        // console.log("------------");
        // console.log(order_date);
        // console.log(c2[date_index]);
        // console.log("------------");
      }
    }
  }

  return cost;
};

var find_prod = function(id){
  var prod = null;
  return prod;
};















