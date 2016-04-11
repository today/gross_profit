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

  // 充当数据源的文件 的文件名 的关键字
  vm.src_files_flag.push("销售订单明细");
  vm.src_files_flag.push("物料清单");
  vm.src_files_flag.push("SCM客户明细");
  vm.src_files_flag.push("内控成本价格变动");
  //vm.src_files_flag.push("xxxxxxx");
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
  vm.base_dir = path.dirname(temp_path ) + "/" ;
  console.log("base_dir: " + vm.base_dir );

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
  return run_flag;
};

var load_order_detail_140 = function(){
  var flag = true;
  var flag_temp = true;

  var obj = null;
  obj = xlsx.parse( vm.base_dir + vm.src_files['销售订单明细'] ); // 读入xlsx文件
  console.log(vm.src_files['销售订单明细']);
  console.log(obj);
  // 取出第一个sheet
  ORDER_DETAIL = obj[0].data;

  // 清洗数据：去除空格
  trim_array_element(ORDER_DETAIL[0]); 
  console.log(ORDER_DETAIL[0]);

  // 检查列名是否正确
  var row_1 = ORDER_DETAIL[0];
  var must_title = [];
  must_title.push("实际交货数量");
  must_title.push("物料描述");
  must_title.push("交货完成状态");
  must_title.push("销售价格");
  must_title.push("订单金额");
  must_title.push("渠道");
  must_title.push("客户");
  must_title.push("客户名称");
  must_title.push("订单数量");
  must_title.push("物料编码");
  must_title.push("物料组");
  must_title.push("物料组描述");
  must_title.push("库存地点");
  must_title.push("实际交货日期");
  must_title.push("城市");
  must_title.push("客户参考号");
  must_title.push("创建日期");
  MSG.put( "开始检查「销售订单明细」文件的标题栏。");
  flag_temp = check_must_title(row_1, must_title);
  flag = flag && flag_temp;

  obj = xlsx.parse( vm.base_dir + vm.src_files['物料清单'] );
  // 取出第一个sheet
  var sheet_1 = obj[0].data;
  var row_1 = sheet_1[0];
  var must_title = [];
  must_title.push("物料编码");
  must_title.push("预期成本价格");
  must_title.push("内控成本价格");
  //must_title.push("开始变动日期");
  must_title.push("产品经理");

  MSG.put( "开始检查「物料清单」文件的标题栏。");
  flag_temp = check_must_title(row_1, must_title);
  flag = flag && flag_temp;

  obj = xlsx.parse( vm.base_dir + vm.src_files['SCM客户明细'] );
  // 取出第一个sheet
  var sheet_1 = obj[0].data;
  var row_1 = sheet_1[0];
  var must_title = [];
  must_title.push("客户");
  must_title.push("地市");
  must_title.push("渠道经理");

  MSG.put( "开始检查「SCM客户明细」文件的标题栏。");
  flag_temp = check_must_title(row_1, must_title);
  flag = flag && flag_temp;

  obj = xlsx.parse( vm.base_dir + vm.src_files['内控成本价格变动'] );
  // 取出第一个sheet
  var sheet_1 = obj[0].data;
  var row_1 = sheet_1[0];
  var must_title = [];
  must_title.push("物料编码");
  must_title.push("预期成本价格");
  must_title.push("内控成本价格");
  must_title.push("开始变动日期");
  MSG.put( "开始检查「内控成本价格变动」文件的标题栏。");
  flag_temp = check_must_title(row_1, must_title);
  flag = flag && flag_temp;
  //MSG.put( "销售订单明细   数据读入成功。");

  return flag;;
};



var copy_order_detail_150 = function(){
  
  var must_col_title = [];
  must_col_title.push(make_title("实际交货数量"));
  must_col_title.push(make_title("物料描述"));
  must_col_title.push(make_title("交货完成状态"));
  must_col_title.push(make_title("销售价格"));
  must_col_title.push(make_title("订单金额"));
  must_col_title.push(make_title("订单类型"));
  must_col_title.push(make_title("渠道"));
  must_col_title.push(make_title("客户"));
  must_col_title.push(make_title("客户名称"));
  must_col_title.push(make_title("订单数量"));
  must_col_title.push(make_title("物料编码"));
  must_col_title.push(make_title("物料组"));
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
  //console.log(must_col_title);

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

  // 进行数据清洗，物料编码，把18位的编码缩减到16位。客户编码，删除前面的两个零。
  var prod_id_index = find_title_index(must_col_title, "物料编码");
  var custom_id_index = find_title_index(must_col_title, "客户");

  if( -1 === prod_id_index ){
    ERR_MSG.put("数据出错：。无法找到「物料编码」列，请检查数据。" );
    return false;
  }
  if( -1 === custom_id_index ){
    ERR_MSG.put("数据出错：。无法找到「客户」列，请检查数据。" );
    return false;
  }

  for(var i=1;i<ORDER_DETAIL_SMALL.length; i++){
    var prod_id_temp = ORDER_DETAIL_SMALL[i][prod_id_index];
    var custom_id_temp = ORDER_DETAIL_SMALL[i][custom_id_index];

    if( 18 === prod_id_temp.length ){
      ORDER_DETAIL_SMALL[i][prod_id_index] = prod_id_temp.substring(2);
      ORDER_DETAIL_SMALL[i][custom_id_index] = custom_id_temp.substring(2);
    }else{
      ERR_MSG.put("数据出错：订单表中的物料编码长度不是18。行数：" + i + " 物料编码：" + prod_id_temp );
      console.log(prod_id_temp);
    }

    ORDER_DETAIL_SMALL[i][custom_id_index] = custom_id_temp.substring(2);
    
  }

  var buffer = xlsx.build([{name: "销售订单明细", data: ORDER_DETAIL_SMALL}]);
  fs.writeFileSync( vm.base_dir + "中间文件_销售数据.xlsx", buffer);

  
  return true;
};

// 第六步：补充数据到工作文件。
// 成本价格取数的时候  应该使用“内控成本价格” 
var fill_field_160 = function(){

  // 获得物料数据。
  var prod_info = getProd_info();
  // var buffer = xlsx.build([{name: "debug物料", data: prod_info}]);
  // fs.writeFileSync(  vm.base_dir + "debug物料.xlsx", buffer);
  
  // 装入  [销售订单明细]
  var obj_sheet = xlsx.parse( vm.base_dir + "中间文件_销售数据.xlsx");
  var order_info =  obj_sheet[0].data;
  MSG.put( " 中间文件_销售数据.XLSX  数据读入成功。");


  // 销售订单表
  var title_array_3 = order_info[0];
  console.log(title_array_3);
  var index_prod_id = find_title_index(title_array_3, "物料编码");
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
    
    if( undefined === cost ){
      ERR_MSG.put("数据出错：物料表中的成本价格未填写。行数：\t" + (i+1) + "\t物料编码：\t" + prod_id_in_order );
      a_order[index_cost] = -1;
    }
    else if( 0 === cost ){
      //ERR_MSG.put("数据出错：物料表中的成本价为 0 。行数：\t" + (i+1) + "\t物料编码：\t" + prod_id_in_order );
      a_order[index_cost] = 0;
    }
    else if( -2 === cost ){
      ERR_MSG.put("数据出错：无法在物料表中找到这个物料。行数：\t" + (i+1) + "\t物料编码：\t" + prod_id_in_order );
      a_order[index_cost] = -2;
    }
    else if( -3 === cost ){
      ERR_MSG.put("数据出错：无法在物料表中找到这个物料。行数：\t" + (i+1) + "\t物料编码：\t" + prod_id_in_order );
      a_order[index_cost] = -3;
    }
    else{
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
  
  var index_prod_id = find_title_index(title_array, "物料编码");
  var index_order_amount = find_title_index(title_array, "订单金额"); 
  var index_count = find_title_index(title_array, "订单数量");
  var index_cost = find_title_index(title_array, "内控成本价格");
  var index_income = find_title_index(title_array, "销售收入");
  var index_gross = find_title_index(title_array, "毛利");
  var index_gross_rate = find_title_index(title_array, "毛利率");


  // 因为要跳过title，所以下标从 1 开始。
  for(var i=1; i<order_info.length; i++){
    var a_order = order_info[i];
    
    var prod_id = a_order[index_prod_id];
    var amount = a_order[index_order_amount];
    var order_count = a_order[index_count];
    var cost = a_order[index_cost];

    if( 0 < cost ){
      var cost_sum = cost * order_count / TAX_RATE;
      var income_sum = amount / TAX_RATE
      a_order[index_income] = income_sum;
      a_order[index_gross] = (income_sum - cost_sum) ;
      a_order[index_gross_rate] = a_order[index_gross] / cost_sum * 100;
    }else{
      console.log("数据  有可能  有错。 成本价=" + cost );
      // ERR_MSG.put("数据出错：成本价数据异常。行数：" + (i+1) + " 成本价：" + cost + " 物料编码：" + prod_id );
    }

  }

  var buffer = xlsx.build([{name: "计算结果_销售毛利", data: order_info}]);
  fs.writeFileSync(  vm.base_dir + "计算结果_销售毛利("+ vm.cost_type +").xlsx", buffer);

  return true;;
};

// 第六步：加总物料毛利。
var calc_prod_180 = function(){
  
  var obj_sheet = xlsx.parse( vm.base_dir + "计算结果_销售毛利("+ vm.cost_type +").xlsx");
  var gross_info =  obj_sheet[0].data;
  MSG.put( " 计算结果_销售毛利.xlsx  数据读入成功。");

  // 获取title
  var title_array = gross_info[0];
  // 结果数组增加字段
  title_array.push("单物料毛利");

  var index_count = find_title_index(title_array, "实际交货数量");
  var index_prod_id = find_title_index(title_array, "物料编码");
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
  fs.writeFileSync(  vm.base_dir + "计算结果_物料毛利("+ vm.cost_type +").xlsx", buffer);

  return true;
};

// 第六步：加总物料组毛利。
var calc_group_190 = function(){
  
  MSG.put( "数据较多，载入约需5秒。请耐心等待。");

  var obj_sheet = xlsx.parse( vm.base_dir + "计算结果_物料毛利("+ vm.cost_type +").xlsx");
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
  index_will_delete.push( find_title_index(title_array, "物料编码") );
  index_will_delete.push( find_title_index(title_array, "库存地点") );
  index_will_delete.push( find_title_index(title_array, "成本单价") );
  index_will_delete.push( find_title_index(title_array, "单物料毛利") );
  // 清除不需要的列
  var thin_gross = del_col_from_array(group_gross_sum, index_will_delete);

  var buffer = xlsx.build([{name: "计算结果_物料组毛利", data: thin_gross}]);
  fs.writeFileSync(  vm.base_dir + "计算结果_物料组毛利("+ vm.cost_type +").xlsx", buffer);

  return true;
};

var getProd_info = function(){
  // 装入  [物料清单.XLSX]
  // 物料清单中编码为 1001000202013110  16位   
  //   销售订单中  001001000202013110  18位
  // 需清洗数据。
  var obj_sheet2 = xlsx.parse( vm.base_dir + vm.src_files['物料清单']);
  //console.log(obj_sheet2);
  var prod_info = obj_sheet2[0].data;
  MSG.put( " 物料清单  数据读入成功。");

  // 补充需要的列
  prod_info[0].push("开始变动日期");

  // 取出必要的列
  var title_array = prod_info[0];
  console.log(title_array);
  var index_must = [];
  index_must.push( find_title_index(title_array, "物料编码") );
  index_must.push( find_title_index(title_array, "物料描述") );
  index_must.push( find_title_index(title_array, "预期成本价格") );
  index_must.push( find_title_index(title_array, "内控成本价格") );
  index_must.push( find_title_index(title_array, "开始变动日期") );
  console.log(index_must);

  // 填充一个默认值到 「开始变动日期」  字段里
  var index_date = find_title_index(title_array, "开始变动日期");
  for(var i=1; i<prod_info.length; i++){
    prod_info[i][index_date] = 40000;
  }
  
  var prod_must_col = select_col_from_array(prod_info, index_must);
  
  // 装入 
  var obj_sheet3 = xlsx.parse( vm.base_dir + vm.src_files['内控成本价格变动']); 
  var price_history = obj_sheet3[0].data;
  MSG.put( " 内控成本价格变动  数据读入成功。");
  // 取出必要的列
  var title_array_2 = price_history[0];

  var index_must_2 = [];
  index_must_2.push( find_title_index(title_array_2, "物料编码") );
  index_must_2.push( find_title_index(title_array_2, "物料描述") );
  index_must_2.push( find_title_index(title_array_2, "预期成本价格") );
  index_must_2.push( find_title_index(title_array_2, "内控成本价格") );
  index_must_2.push( find_title_index(title_array_2, "开始变动日期") );
  console.log(index_must_2);
  var cost_history_col = select_col_from_array(price_history, index_must_2);

  // 合并价格变动数组到物料数组中
  for(var i=1; i<cost_history_col.length; i++){
    prod_must_col.push(cost_history_col[i]);
  }

  return prod_must_col;
}

var fill_branch_200 = function(){

  var obj_sheet = xlsx.parse( vm.base_dir + "计算结果_销售毛利("+ vm.cost_type +").xlsx");
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
  // 第四步，1001仓出库的 按照客户编码区分  1开头的就是自有渠道 2开头的就是零售渠道 8开头的就是分销渠道
  // 第五步，仓库是 任意 , 客户编码是 8 开头，是分销渠道。
  // 第六步，如果仍然有剩余数据，报错。

  // 注：销售表里的 客户编码 也需要做数据清洗，已经在前面做了。
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
  for(var i=0; i<temp_array.length; i++){
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
  for(var i=0; i<temp_array2.length; i++){
    var temp_order = temp_array2[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    if( '002'===custom_id.substring(0,3) ){
      temp_order[index_branch] = '零售渠道';
    }else{
      temp_array3.push(temp_order);
    }
  }

  var temp_array4 = [];
  for(var i=0; i<temp_array3.length; i++){
    var temp_order = temp_array3[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    if( '1001' === temp_id.substring(0,4) ){

      if( '001'===custom_id.substring(0,1) ){
        temp_order[index_branch] = '自有渠道';
      }else if( '002'===custom_id.substring(0,1) ){
        temp_order[index_branch] = '零售渠道';
      }else if( '008'===custom_id.substring(0,1) ){
        temp_order[index_branch] = '分销渠道';
      }else{
        temp_array4.push(temp_order);
      }

    }else{
      temp_array4.push(temp_order);
    }
  }

  var temp_array5 = [];
  for(var i=0; i<temp_array4.length; i++){
    var temp_order = temp_array4[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    
    if( '8'===custom_id.substring(0,1) || '001'===custom_id.substring(0,3) ){
      temp_order[index_branch] = '分销渠道';
    }else{
      temp_array5.push(temp_order);
    }
  }

  for(var i=0; i<temp_array5.length; i++){
    var temp_order = temp_array5[i];
    var temp_id = temp_order[index_warehouse_id];
    var custom_id = temp_order[index_custom_id];
    temp_order[index_branch] = '未确定渠道';

    ERR_MSG.put("数据出错：发现无渠道归属的订单。 库存地点=" + temp_id + " 客户编码=" +　custom_id );
    //console.log(temp_order);
    console.log("库存地点=" + temp_id + " 客户编码=" + custom_id);
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
  }

  // P001  西安城区铺货仓
  // P002  咸阳铺货仓
  // P003  宝鸡铺货仓
  // P004  渭南铺货仓
  // P005  铜川铺货仓
  // P006  延安铺货仓
  // P007  榆林铺货仓
  // P008  汉中铺货仓
  // P009  安康铺货仓
  // P010  商洛铺货仓
  // P011  西安郊县铺货仓
  var get_self_branch_city = function( branch ){
    var city = "出错";
    if( "" === branch ) city = "";
    else if ( "P001" === branch ) city = "西安城区";
    else if ( "P002" === branch ) city = "咸阳";
    else if ( "P003" === branch ) city = "宝鸡";
    else if ( "P004" === branch ) city = "渭南";
    else if ( "P005" === branch ) city = "铜川";
    else if ( "P006" === branch ) city = "延安";
    else if ( "P007" === branch ) city = "榆林";
    else if ( "P008" === branch ) city = "汉中";
    else if ( "P009" === branch ) city = "安康";
    else if ( "P010" === branch ) city = "商洛";
    else if ( "P011" === branch ) city = "西安郊县";
    else if ( "P099" === branch ) city = "零售中心";
    else{
      city = "出错";
      console.log("客户的地区归属出错。 branch = #" + branch + "#");
    } 

    return city;
  };


  // 单独处理 自有渠道 的 地市归属
  var index_warehouse_id = find_title_index(title_array, "库存地点");
  var index_branch = find_title_index(title_array, "渠道归属");
  for(var i=1; i<gross_info.length; i++ ){
    var order = gross_info[i];
    var branch = order[index_branch];
    if( "自有渠道" === branch ){
      var warehouse_id = order[index_warehouse_id];
      var city = get_self_branch_city(warehouse_id);
      //console.log(city);
      order[index_city] = city;
    }
    //if(i>10)break;
  }

  var buffer = xlsx.build([{name: "销售数据(渠道归属和地市归属)", data: gross_info}]);
  fs.writeFileSync(  vm.base_dir + "中间文件_销售数据(渠道归属和地市归属).xlsx", buffer);
  
  return true;
};

// TODO   del  next   line
temp_summary = 0

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
  title_array_dest.push("自有渠道毛利率");
  title_array_dest.push("自有渠道收入占比");

  title_array_dest.push("分销渠道销量");  
  title_array_dest.push("分销渠道收入");  
  title_array_dest.push("分销渠道毛利"); 
  title_array_dest.push("分销渠道毛利率");
  title_array_dest.push("分销渠道收入占比");

  title_array_dest.push("电子渠道销量");  
  title_array_dest.push("电子渠道收入");  
  title_array_dest.push("电子渠道毛利"); 
  title_array_dest.push("电子渠道毛利率");
  title_array_dest.push("电子渠道收入占比");

  title_array_dest.push("零售渠道销量");  
  title_array_dest.push("零售渠道收入");  
  title_array_dest.push("零售渠道毛利"); 
  title_array_dest.push("零售渠道毛利率");
  title_array_dest.push("零售渠道收入占比");

  title_array_dest.push("未确定渠道销量");  
  title_array_dest.push("未确定渠道收入");  
  title_array_dest.push("未确定渠道毛利"); 
  title_array_dest.push("未确定渠道毛利率");
  title_array_dest.push("未确定渠道收入占比");

  title_array_dest.push("合计销量");  
  title_array_dest.push("合计收入");  
  title_array_dest.push("合计毛利");
  title_array_dest.push("合计毛利率");
  title_array_dest.push("合计收入占比");

  var make_summary_line = function(title){ 
    var temp_array = make_array(30, 0);
    var data = [""].concat(temp_array);
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
  data_array["零售中心"] = make_summary_line('零售中心');
  data_array["其他"]    = make_summary_line('其他');
  data_array["数据错误"] = make_summary_line('数据错误');
  data_array["合计"]    = make_summary_line('合计');

  // 加总数量，收入，毛利，毛利率  的函数。
  var calc_summary = function( order, data){
    var count = order[index_count];
    var income = order[index_income];
    var gross = order[index_gross];
    var gross_rate = 0;

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
      offset = 5;
    }else if ( "电子渠道" === branch ){
      offset = 10;
    }else if ( "零售渠道" === branch ){
      offset = 15;
    }else if ( "未确定渠道" === branch ){
      offset = 20;
    }else {
      console.log("渠道归属 为空");
    }

    // temp_summary += count;
    // console.log("temp_summary=" + temp_summary );

    data[offset+1] += count;
    data[offset+2] += income;
    data[offset+3] += gross;
    if( data[offset+2] != 0 ){
      data[offset+4] = data[offset+3]/data[offset+2];
    }

    data[26] += count;
    data[27] += income;
    data[28] += gross;
    if( data[27] != 0 ){
      data[29] = data[28]/data[27];
    }

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
  full_data.push(data_array["零售中心"] );
  full_data.push(data_array["其他"]    );
  full_data.push(data_array["数据错误"]    );
  full_data.push(data_array["合计"] );

  //  计算列合计
  var temp_hj = data_array['合计'];
  for(var i=1; i<temp_hj.length; i++ ){
    var temp_one_col = select_one_col_from_table(full_data, i);
    //console.log(temp_one_col);
    temp_hj[i] = array_sum(temp_one_col);
  }

  //  计算渠道收入占比。
  var index_zy = find_title_index(title_array_dest, "自有渠道收入占比");
  var index_fx = find_title_index(title_array_dest, "分销渠道收入占比");
  var index_dz = find_title_index(title_array_dest, "电子渠道收入占比");
  var index_ls = find_title_index(title_array_dest, "零售渠道收入占比");
  var index_wqd = find_title_index(title_array_dest, "未确定渠道收入占比");
  var index_hjsr = find_title_index(title_array_dest, "合计收入");
  for( var i=1; i<full_data.length; i++ ){
    var temp_data = full_data[i];
    temp_data[index_zy] = temp_data[index_zy-3] / temp_data[index_hjsr];
    temp_data[index_fx] = temp_data[index_fx-3] / temp_data[index_hjsr];
    temp_data[index_dz] = temp_data[index_dz-3] / temp_data[index_hjsr];
    temp_data[index_ls] = temp_data[index_ls-3] / temp_data[index_hjsr];
    temp_data[index_wqd] = temp_data[index_wqd-3] / temp_data[index_hjsr];
  }
  //  计算 最后一列  地市收入占比。
  var index_hj = find_title_index(title_array_dest, "合计收入占比");
  var array_hj = select_one_col_from_table(full_data, index_hj);
  //console.log(array_hj);
  for( var i=1; i<array_hj.length-1; i++ ){
    console.log(full_data[i][index_hj-3]);
    console.log(full_data[full_data.length-1][index_hjsr]);
    full_data[i][index_hj] = full_data[i][index_hj-3] / full_data[full_data.length-1][index_hjsr];
  }


  // 地市渠道汇总
  var obj_summary = {name: "地市渠道销量收入毛利汇总", data: full_data};
  //console.log(obj_summary);
  // 地市汇总表  
  var city_summary = select_col_from_array(full_data,[0,27,28,29,30]);
  city_summary = add_col_for_table(city_summary, "");
  city_summary[0][5] = "备注";
  city_summary.splice(0, 0, ["地市汇总表",""]);
  city_summary.push( ["合计",""]);
  city_summary.push( ["毛利率",""]);
  var obj_city = {name: "地市汇总表", data: city_summary};
  console.log(obj_city);

  // 渠道汇总表 
  var branch_summary = select_col_from_array(full_data,[0,27,28,29,30]);
  var obj_branch = {name: "渠道汇总表", data: branch_summary};
  console.log(obj_branch);

  var sheets = [];
  sheets.push(obj_summary);
  sheets.push(obj_city);
  sheets.push(obj_branch);
  
  var buffer = xlsx.build(sheets);
  fs.writeFileSync(  vm.base_dir + "计算结果_地市渠道销量收入毛利汇总表("+ vm.cost_type +").xlsx", buffer);
  return true;
};

// 第十二步：计算渠道经理，产品经理，客户 毛利贡献。
var gross_contribute_230 = function () {

  return true;
}

var format_240 = function () {
  // 处理最终结果，把数字格式化成两位小数
  var file_list = [];
  file_list.push("计算结果_销售毛利("+ vm.cost_type +")");
  file_list.push("计算结果_物料毛利("+ vm.cost_type +")");
  file_list.push("计算结果_物料组毛利("+ vm.cost_type +")");
  file_list.push("计算结果_地市渠道销量收入毛利汇总表("+ vm.cost_type +")");

  for(var i=0; i<file_list.length; i++){
    var obj_sheet = xlsx.parse( vm.base_dir + file_list[i] + ".xlsx");

    for(var j=0; j<obj_sheet.length; j++ ){
      var data_array =  obj_sheet[j].data;
      var full_data = format_data( data_array );
      obj_sheet[j].data = full_data;
    }

    var buffer = xlsx.build(obj_sheet);
    fs.writeFileSync(  vm.base_dir + file_list[i] + ".xlsx", buffer);
  }

  return true;
}

var finish_250 = function () {
  // 打开计算结果的文件夹
  var os = require('os');
  var os_name = os.platform();
  var exec = require('child_process').exec; 

  if( "darwin" === os_name ){
    
    var cmdStr = 'open ' + vm.base_dir ;
    console.log(cmdStr);
    exec(cmdStr, function(err,stdout,stderr){
        if(err) {
            console.log('error:'+stderr);
        } 
    });
  }
  else if( "win32" === os_name ){
    var cmdStr = 'start ' + vm.base_dir ;
    exec(cmdStr, function(err,stdout,stderr){
        if(err) {
            console.log('error:' + stderr);
        } 
    });
  }else{
    console.log("暂不支持的操作系统： " + os_name );
  }
  return true;
}

/***************  分隔线  ******************/
var format_percent = function(a_num){
  return ( a_num * 100 ).toFixed(2) + "%";
}

var format_two_decimal = function(a_num){
  return a_num.toFixed(2);
}

var format_data = function(table_array){
  for(var i=0;i<table_array.length; i++){
    var temp_array = table_array[i];
    //console.log(temp_array);
    for(var j=0;j<temp_array.length; j++){
      var temp_cell = temp_array[j];
      if( typeof(temp_cell) === typeof(3) ){   // 如果是数字
        if( 1 > temp_cell && -1 < temp_cell ) {  // 如果小于1，就处理成百分比格式
          temp_array[j] = format_percent( temp_cell );
        }
        // 如果不是整数 并且是NaN，也就是数字
        else if( (! Number.isInteger(temp_cell) ) && (! Number.isNaN(temp_cell) ) ){   
          temp_array[j] = format_two_decimal( temp_cell );
        }
      }
    }
    //console.log(temp_array);
  }
  return table_array;
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



var getCity = function(custom_info, custom_id){

  var title_array = custom_info[0];
  //console.log(title_array);
  custom_no_index = find_title_index(title_array, "客户");
  city_index = find_title_index(title_array, "地市");
  //console.log(custom_no_index);
  //console.log(city_index);
  var ret_city = null;

  var no = "xxx";
  var no2 = "yyy";

  if( typeof(custom_id) === typeof(123) ){
    no = custom_id;
  }else{
    no = parseInt( custom_id );
  }

  for(var i=1; i<custom_info.length; i++){
    var custom = custom_info[i];
    no2 = custom[custom_no_index];
    
    var city = custom[city_index];
    if( city ){
      city = custom[city_index].trim();
    }else{
      console.log(custom[custom_no_index]);
      console.log(custom[city_index]);
      ERR_MSG.put("取客户地市出错。 序号: #"+ i + "#   地市 : #" + city + "#" );
      continue;
      // 
    } 

    if( typeof(no2) === typeof(123) ){
      no2 = no2;
    }else{
      no2 = parseInt(no2);
    }

    //console.log(custom_id);
    //console.log(no);
    if( no === no2 ){
      if( city.indexOf("西安城区") > -1 ){
        ret_city = "西安城区";
      }
      else if( city.indexOf("西安郊县") > -1 ){
        ret_city = "西安郊县";
      }
      else if( city.indexOf("西安") > -1 ){
        ret_city = "西安城区";
      }
      else if( city.indexOf("咸阳") > -1 ){
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
        ERR_MSG.put("客户地区归属出错。 地区 : #" + no + "#" );
        console.log("客户的地区归属出错。 #" + no + "# #" + no2 + "# #" + city + "#");
      }
      break;
    }
  }

  if( ret_city == null ){
    console.log("未能找到对应的客户。 客户编码: #" + no + "# #" + ret_city + "#");
    ERR_MSG.put("未能找到对应的客户。 客户编码: #" + no + "# #" + ret_city + "#" );
    //console.log("#" + typeof(no) + "# #" + typeof(no2) + "#");
  }

  return ret_city;
}

var getCost = function(prod_info, id, order_date){
  //console.log(id);
  var temp_array = [];
  var temp_index = -1;   // 貌似这个变量没有用到。
  var prod_id = null;
  var id_a = id.trim();

  var title_array = prod_info[0];
  id_index = find_title_index(title_array, "物料编码");
  //cost_index = find_title_index(title_array, "预期成本价格");
  //cost_index = find_title_index(title_array, "内控成本价格");
  cost_index = find_title_index(title_array, vm.cost_type );
  
  date_index = find_title_index(title_array, "开始变动日期");

  for(var i=1; i<prod_info.length; i++){
    prod_id = prod_info[i][id_index];
    // 把数字转为字符串
    if(_.isNumber(prod_id)){   
      prod_id = ""+prod_id;
    }

    if( isblank(prod_id) ){
      temp_index = -1;
    }else{
      var id_b = prod_id.trim(); 

      if( id_a == id_b ){
        temp_array.push(prod_info[i]);
      }
    }
  }

  // 看看哪个价钱是当前的价钱
  var cost = -3;
  if( temp_array.length === 0 ){
    console.log("Warning: cost not found. prod_id=" +　id_a　);
    cost = -2;
  }
  else if( temp_array.length === 1 ){
    cost = temp_array[0][cost_index];
    // if( undefined === cost ) {
    //   console.log("11111111111");
    //   console.log(temp_array);
    //   console.log(cost_index);
    // }
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
      // 
      if( order_date >= c1[date_index] && order_date < c2[date_index] ){
        cost = c1[cost_index];
        // if( undefined === cost ){
        //   console.log("222222222222");
        //   console.log(c1);
        //   console.log(cost_index);
        // }
        // console.log("------------");
        // console.log(c1[date_index]);
        // console.log(order_date);
        // console.log(c2[date_index]);
        // console.log("------------");
        break;
      }
      else if( i === (temp_array.length-1) && order_date >= c2[date_index]){
        cost = c2[cost_index];
        if( undefined === cost ) console.log("cost is undefined");
        // console.log("------------");
        // console.log(order_date);
        // console.log(c2[date_index]);
        // console.log("------------");
      }
    }
  }

  return cost;
};


var check_must_title = function(target, keywords ){
  var ret_flag = false;
  for(var i=0; i<keywords.length; i++ ){
    var kw = keywords[i];
    console.log(kw);
    if( -1 === _.indexOf(target, kw)){
      // 显示出错提示。
      ERR_MSG.put("数据出错：。无法找到「"+ kw +"」列，请检查数据的第一行。" );
      ret_flag = true;
    }
  }
  return ret_flag;
}













