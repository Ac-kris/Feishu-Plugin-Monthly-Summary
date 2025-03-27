import { 
  bitable, 
  FieldType, 
  ITable, 
  ICurrencyField, 
  INumberField, 
  ISingleSelectField, 
  IDateTimeField,
  NumberFormatter,
  CurrencyCode
} from '@lark-base-open/js-sdk';

/**
 * 月度制备费汇总表生成器
 */
export class MonthlyReportGenerator {
  // 源表字段名称
  private readonly SOURCE_FIELD_NAMES = {
    PRODUCT_CATEGORY: '产品类别',
    OUT_DATE: '申请出库日期',
    SHARES: '份数',
    PREPARATION_FEE: '制备费',
    REWARD_DEDUCTION: '出库订单（多单）奖励扣减',
    CANCEL_FEE: '制剂取消收取费用',
    URGENT_FEE: '制剂加急/非正常出库时间收取费用'
  };

  // 目标表字段名称
  private readonly TARGET_FIELD_NAMES = {
    PROJECT: '项目',
    PERSON_COUNT: '人数',
    SHARES: '份数',
    PREPARATION_FEE: '制备费',
    REWARD_DEDUCTION: '出库订单奖励扣减',
    CANCEL_FEE: '制剂取消收取费用',
    URGENT_FEE: '制剂加急/非正常出库时间收取费用'
  };

  /**
   * 生成月度汇总表
   * @param sourceTableId 源表ID
   * @param month 月份(1-12)
   */
  public async generate(sourceTableId: string, month: number): Promise<void> {
    if (!sourceTableId) {
      throw new Error('源表ID不能为空');
    }
    
    if (month < 1 || month > 12) {
      throw new Error('月份必须在1-12之间');
    }
    
    try {
      // 获取源表
      const sourceTable = await bitable.base.getTableById(sourceTableId);
      if (!sourceTable) {
        throw new Error(`未找到ID为 ${sourceTableId} 的表格`);
      }
      
      // 创建新表
      const monthStr = String(month);
      const currentDate = new Date();
      const year = currentDate.getFullYear();
      const newTableName = `${year}年${monthStr}月支付细胞工程制剂制备费`;
      
      console.log(`正在创建新表: ${newTableName}`);
      
      // 修复addTable的参数类型，添加必要的fields属性
      const newTableResult = await bitable.base.addTable({ 
        name: newTableName,
        fields: []
      });
      
      if (!newTableResult || !newTableResult.tableId) {
        throw new Error('创建新表失败');
      }
      
      // 获取新创建的表
      const newTable = await bitable.base.getTableById(newTableResult.tableId);
      if (!newTable) {
        throw new Error(`无法访问新创建的表格: ${newTableResult.tableId}`);
      }
      
      console.log('创建字段中...');
      // 创建字段
      await this.createFields(newTable);
      
      console.log('获取源表数据中...');
      // 获取源表中的数据
      const sourceData = await this.getSourceData(sourceTable, month);
      
      // 检查是否找到符合条件的数据
      if (Object.keys(sourceData).length === 0) {
        throw new Error(`未找到${month}月的数据，请确认原表中存在该月份的记录`);
      }
      
      console.log('写入数据到新表中...');
      // 将数据写入新表
      await this.writeDataToNewTable(newTable, sourceData);
      
      console.log('汇总表生成完成');
    } catch (error) {
      console.error('生成汇总表时发生错误:', error);
      if (error instanceof Error) {
        throw error;
      } else {
        throw new Error('生成汇总表时发生未知错误');
      }
    }
  }
  
  /**
   * 创建字段
   */
  private async createFields(table: ITable): Promise<void> {
    try {
      // 项目（下拉选择）
      await table.addField({
        type: FieldType.SingleSelect,
        name: this.TARGET_FIELD_NAMES.PROJECT
      });
      
      // 人数（整数数字）
      await table.addField({
        type: FieldType.Number,
        name: this.TARGET_FIELD_NAMES.PERSON_COUNT,
        property: {
          formatter: NumberFormatter.INTEGER
        }
      });
      
      // 份数（带2位小数的数字）
      await table.addField({
        type: FieldType.Number,
        name: this.TARGET_FIELD_NAMES.SHARES,
        property: {
          formatter: NumberFormatter.DIGITAL_ROUNDED_2
        }
      });
      
      // 制备费（金额，2位小数）
      await table.addField({
        type: FieldType.Currency,
        name: this.TARGET_FIELD_NAMES.PREPARATION_FEE,
        property: {
          currencyCode: CurrencyCode.CNY
        }
      });
      
      // 出库订单奖励扣减（金额，2位小数）
      await table.addField({
        type: FieldType.Currency,
        name: this.TARGET_FIELD_NAMES.REWARD_DEDUCTION,
        property: {
          currencyCode: CurrencyCode.CNY
        }
      });
      
      // 制剂取消收取费用（金额，2位小数）
      await table.addField({
        type: FieldType.Currency,
        name: this.TARGET_FIELD_NAMES.CANCEL_FEE,
        property: {
          currencyCode: CurrencyCode.CNY
        }
      });
      
      // 制剂加急/非正常出库时间收取费用（金额，2位小数）
      await table.addField({
        type: FieldType.Currency,
        name: this.TARGET_FIELD_NAMES.URGENT_FEE,
        property: {
          currencyCode: CurrencyCode.CNY
        }
      });
    } catch (error) {
      console.error('创建字段时发生错误:', error);
      throw new Error('创建字段失败，请检查字段配置');
    }
  }
  
  /**
   * 获取源表数据
   */
  private async getSourceData(sourceTable: ITable, month: number): Promise<Record<string, any>> {
    try {
      // 获取字段元数据
      const fields = await sourceTable.getFieldMetaList();
      if (!fields || fields.length === 0) {
        throw new Error('源表中没有找到任何字段');
      }
      
      // 获取字段ID
      const productCategoryFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.PRODUCT_CATEGORY)?.id;
      const outDateFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.OUT_DATE)?.id;
      const sharesFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.SHARES)?.id;
      const preparationFeeFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.PREPARATION_FEE)?.id;
      const rewardDeductionFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.REWARD_DEDUCTION)?.id;
      const cancelFeeFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.CANCEL_FEE)?.id;
      const urgentFeeFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.URGENT_FEE)?.id;
      
      // 验证缺失的字段
      const missingFields = [];
      if (!productCategoryFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.PRODUCT_CATEGORY);
      if (!outDateFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.OUT_DATE);
      if (!sharesFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.SHARES);
      if (!preparationFeeFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.PREPARATION_FEE);
      if (!rewardDeductionFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.REWARD_DEDUCTION);
      if (!cancelFeeFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.CANCEL_FEE);
      if (!urgentFeeFieldId) missingFields.push(this.SOURCE_FIELD_NAMES.URGENT_FEE);
      
      if (missingFields.length > 0) {
        throw new Error(`源表中缺少以下必要的字段: ${missingFields.join(', ')}`);
      }
      
      // 获取字段对象 - 使用非空断言操作符，因为上面已经验证了字段ID非空
      const productCategoryField = await sourceTable.getField<ISingleSelectField>(productCategoryFieldId!);
      const outDateField = await sourceTable.getField<IDateTimeField>(outDateFieldId!);
      const sharesField = await sourceTable.getField<INumberField>(sharesFieldId!);
      const preparationFeeField = await sourceTable.getField<ICurrencyField>(preparationFeeFieldId!);
      const rewardDeductionField = await sourceTable.getField<ICurrencyField>(rewardDeductionFieldId!);
      const cancelFeeField = await sourceTable.getField<ICurrencyField>(cancelFeeFieldId!);
      const urgentFeeField = await sourceTable.getField<ICurrencyField>(urgentFeeFieldId!);
      
      // 获取记录ID列表
      const recordIdList = await sourceTable.getRecordIdList();
      if (!recordIdList || recordIdList.length === 0) {
        throw new Error('源表中没有找到任何记录');
      }
      
      console.log(`共找到 ${recordIdList.length} 条记录，准备处理...`);
      
      // 分组数据
      const groupedData: Record<string, {
        count: number;
        shares: number;
        preparationFee: number;
        rewardDeduction: number;
        cancelFee: number;
        urgentFee: number;
      }> = {};
      
      let processedRecordCount = 0;
      let matchedRecordCount = 0;
      
      // 遍历记录
      for (const recordId of recordIdList) {
        try {
          // 获取日期值
          const outDateValue = await outDateField.getValue(recordId);
          
          // 检查是否有日期值
          if (!outDateValue) continue;
          
          // 解析日期
          let outDate: Date;
          try {
            outDate = new Date(outDateValue);
            
            // 检查日期是否有效
            if (isNaN(outDate.getTime())) {
              console.warn(`记录 ${recordId} 的日期格式无效: ${outDateValue}`);
              continue;
            }
          } catch (error) {
            console.warn(`解析记录 ${recordId} 的日期时出错: ${outDateValue}`);
            continue;
          }
          
          const outDateMonth = outDate.getMonth() + 1; // 月份从0开始，需要+1
          
          // 筛选指定月份的记录
          if (outDateMonth !== month) continue;
          
          matchedRecordCount++;
          
          // 获取产品类别
          const productCategory = await productCategoryField.getValue(recordId);
          
          // 如果没有产品类别则跳过
          if (!productCategory) {
            console.warn(`记录 ${recordId} 没有产品类别，已跳过`);
            continue;
          }
          
          // 将单选值转换为字符串用作键
          const categoryKey = String(productCategory);
          
          // 初始化分组数据
          if (!groupedData[categoryKey]) {
            groupedData[categoryKey] = {
              count: 0,
              shares: 0,
              preparationFee: 0,
              rewardDeduction: 0,
              cancelFee: 0,
              urgentFee: 0
            };
          }
          
          // 累加计数
          groupedData[categoryKey].count += 1;
          
          // 安全获取数值并累加
          const safeGetNumberValue = async (field: INumberField | ICurrencyField, recordId: string): Promise<number> => {
            try {
              const value = await field.getValue(recordId);
              if (value === null || value === undefined) return 0;
              const numValue = Number(value);
              return isNaN(numValue) ? 0 : numValue;
            } catch (error) {
              console.warn(`获取字段值失败: ${error}`);
              return 0;
            }
          };
          
          // 累加份数
          groupedData[categoryKey].shares += await safeGetNumberValue(sharesField, recordId);
          
          // 累加制备费
          groupedData[categoryKey].preparationFee += await safeGetNumberValue(preparationFeeField, recordId);
          
          // 累加出库订单奖励扣减
          groupedData[categoryKey].rewardDeduction += await safeGetNumberValue(rewardDeductionField, recordId);
          
          // 累加制剂取消收取费用
          groupedData[categoryKey].cancelFee += await safeGetNumberValue(cancelFeeField, recordId);
          
          // 累加制剂加急/非正常出库时间收取费用
          groupedData[categoryKey].urgentFee += await safeGetNumberValue(urgentFeeField, recordId);
          
          processedRecordCount++;
        } catch (error) {
          console.warn(`处理记录 ${recordId} 时发生错误:`, error);
          // 继续处理下一条记录，不中断整个处理流程
        }
      }
      
      console.log(`处理完成。共处理 ${processedRecordCount} 条记录，${matchedRecordCount} 条匹配指定月份，得到 ${Object.keys(groupedData).length} 个产品类别分组`);
      
      return groupedData;
    } catch (error) {
      console.error('获取源表数据时发生错误:', error);
      if (error instanceof Error) {
        throw error;
      } else {
        throw new Error('获取源表数据时发生未知错误');
      }
    }
  }
  
  /**
   * 将数据写入新表
   */
  private async writeDataToNewTable(newTable: ITable, groupedData: Record<string, any>): Promise<void> {
    try {
      // 获取字段元数据
      const fields = await newTable.getFieldMetaList();
      if (!fields || fields.length === 0) {
        throw new Error('新表中没有找到任何字段');
      }
      
      // 获取字段ID
      const projectFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.PROJECT)?.id;
      const personCountFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.PERSON_COUNT)?.id;
      const sharesFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.SHARES)?.id;
      const preparationFeeFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.PREPARATION_FEE)?.id;
      const rewardDeductionFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.REWARD_DEDUCTION)?.id;
      const cancelFeeFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.CANCEL_FEE)?.id;
      const urgentFeeFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.URGENT_FEE)?.id;
      
      // 验证缺失的字段
      const missingFields = [];
      if (!projectFieldId) missingFields.push(this.TARGET_FIELD_NAMES.PROJECT);
      if (!personCountFieldId) missingFields.push(this.TARGET_FIELD_NAMES.PERSON_COUNT);
      if (!sharesFieldId) missingFields.push(this.TARGET_FIELD_NAMES.SHARES);
      if (!preparationFeeFieldId) missingFields.push(this.TARGET_FIELD_NAMES.PREPARATION_FEE);
      if (!rewardDeductionFieldId) missingFields.push(this.TARGET_FIELD_NAMES.REWARD_DEDUCTION);
      if (!cancelFeeFieldId) missingFields.push(this.TARGET_FIELD_NAMES.CANCEL_FEE);
      if (!urgentFeeFieldId) missingFields.push(this.TARGET_FIELD_NAMES.URGENT_FEE);
      
      if (missingFields.length > 0) {
        throw new Error(`新表字段创建失败，缺少字段: ${missingFields.join(', ')}`);
      }
      
      // 获取字段对象 - 使用非空断言操作符，因为上面已经验证了字段ID非空
      const projectField = await newTable.getField(projectFieldId!);
      const personCountField = await newTable.getField(personCountFieldId!);
      const sharesField = await newTable.getField(sharesFieldId!);
      const preparationFeeField = await newTable.getField(preparationFeeFieldId!);
      const rewardDeductionField = await newTable.getField(rewardDeductionFieldId!);
      const cancelFeeField = await newTable.getField(cancelFeeFieldId!);
      const urgentFeeField = await newTable.getField(urgentFeeFieldId!);
      
      console.log(`开始写入 ${Object.keys(groupedData).length} 条汇总数据...`);
      
      try {
        // 尝试使用批量添加记录的方式（更高效）
        const records = Object.entries(groupedData).map(([productCategory, data]) => {
          // 确保所有字段ID都存在
          if (!projectFieldId || !personCountFieldId || !sharesFieldId || 
              !preparationFeeFieldId || !rewardDeductionFieldId || !cancelFeeFieldId || !urgentFeeFieldId) {
            throw new Error('字段ID不完整，无法创建记录');
          }
            
          return {
            fields: {
              [projectFieldId]: productCategory,
              [personCountFieldId]: data.count,
              [sharesFieldId]: data.shares,
              [preparationFeeFieldId]: data.preparationFee,
              [rewardDeductionFieldId]: data.rewardDeduction,
              [cancelFeeFieldId]: data.cancelFee,
              [urgentFeeFieldId]: data.urgentFee
            }
          };
        });
        
        if (records.length > 0) {
          await newTable.addRecords(records);
          console.log(`成功批量添加 ${records.length} 条记录`);
          return;
        }
      } catch (error) {
        console.warn('批量添加记录失败，将回退到单条添加模式:', error);
        // 如果批量添加失败，继续使用单条添加的方式
      }
      
      // 遍历分组数据，创建记录（单条添加方式）
      let successCount = 0;
      for (const [productCategory, data] of Object.entries(groupedData)) {
        try {
          // 创建新记录
          const recordId = await newTable.addRecord();
          
          // 设置项目（产品类别）
          await projectField.setValue(recordId, productCategory);
          
          // 设置人数
          await personCountField.setValue(recordId, data.count);
          
          // 设置份数（四舍五入保留2位小数）
          await sharesField.setValue(recordId, Math.round(data.shares * 100) / 100);
          
          // 设置制备费
          await preparationFeeField.setValue(recordId, Math.round(data.preparationFee * 100) / 100);
          
          // 设置出库订单奖励扣减
          await rewardDeductionField.setValue(recordId, Math.round(data.rewardDeduction * 100) / 100);
          
          // 设置制剂取消收取费用
          await cancelFeeField.setValue(recordId, Math.round(data.cancelFee * 100) / 100);
          
          // 设置制剂加急/非正常出库时间收取费用
          await urgentFeeField.setValue(recordId, Math.round(data.urgentFee * 100) / 100);
          
          successCount++;
        } catch (error) {
          console.error(`添加记录 ${productCategory} 时发生错误:`, error);
        }
      }
      
      console.log(`成功添加 ${successCount} 条记录，共 ${Object.keys(groupedData).length} 条`);
      
      if (successCount === 0) {
        throw new Error('没有成功写入任何记录');
      } else if (successCount < Object.keys(groupedData).length) {
        console.warn(`部分记录写入失败: ${successCount}/${Object.keys(groupedData).length}`);
      }
    } catch (error) {
      console.error('写入数据到新表时发生错误:', error);
      if (error instanceof Error) {
        throw error;
      } else {
        throw new Error('写入数据到新表时发生未知错误');
      }
    }
  }
} 