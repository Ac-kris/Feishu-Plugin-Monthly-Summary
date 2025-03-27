import { bitable, FieldType, ITable, IField, ICurrencyField, INumberField, ISingleSelectField, IDateTimeField } from '@lark-base-open/js-sdk';

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
    // 获取源表
    const sourceTable = await bitable.base.getTableById(sourceTableId);
    
    // 创建新表
    const monthStr = String(month);
    const newTableName = `${monthStr}月支付细胞工程制剂制备费`;
    const newTable = await bitable.base.addTable({ name: newTableName });
    
    // 创建字段
    await this.createFields(newTable);
    
    // 获取源表中的数据
    const sourceData = await this.getSourceData(sourceTable, month);
    
    // 将数据写入新表
    await this.writeDataToNewTable(newTable, sourceData);
  }
  
  /**
   * 创建字段
   */
  private async createFields(table: ITable): Promise<void> {
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
        formatter: "0"
      }
    });
    
    // 份数（带2位小数的数字）
    await table.addField({
      type: FieldType.Number,
      name: this.TARGET_FIELD_NAMES.SHARES,
      property: {
        formatter: "0.00"
      }
    });
    
    // 制备费（金额，2位小数）
    await table.addField({
      type: FieldType.Currency,
      name: this.TARGET_FIELD_NAMES.PREPARATION_FEE,
      property: {
        currencyCode: "CNY"
      }
    });
    
    // 出库订单奖励扣减（金额，2位小数）
    await table.addField({
      type: FieldType.Currency,
      name: this.TARGET_FIELD_NAMES.REWARD_DEDUCTION,
      property: {
        currencyCode: "CNY"
      }
    });
    
    // 制剂取消收取费用（金额，2位小数）
    await table.addField({
      type: FieldType.Currency,
      name: this.TARGET_FIELD_NAMES.CANCEL_FEE,
      property: {
        currencyCode: "CNY"
      }
    });
    
    // 制剂加急/非正常出库时间收取费用（金额，2位小数）
    await table.addField({
      type: FieldType.Currency,
      name: this.TARGET_FIELD_NAMES.URGENT_FEE,
      property: {
        currencyCode: "CNY"
      }
    });
  }
  
  /**
   * 获取源表数据
   */
  private async getSourceData(sourceTable: ITable, month: number): Promise<Record<string, any>> {
    // 获取字段元数据
    const fields = await sourceTable.getFieldMetaList();
    
    // 获取字段ID
    const productCategoryFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.PRODUCT_CATEGORY)?.id;
    const outDateFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.OUT_DATE)?.id;
    const sharesFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.SHARES)?.id;
    const preparationFeeFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.PREPARATION_FEE)?.id;
    const rewardDeductionFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.REWARD_DEDUCTION)?.id;
    const cancelFeeFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.CANCEL_FEE)?.id;
    const urgentFeeFieldId = fields.find(field => field.name === this.SOURCE_FIELD_NAMES.URGENT_FEE)?.id;
    
    // 验证字段是否存在
    if (!productCategoryFieldId || !outDateFieldId || !sharesFieldId || 
        !preparationFeeFieldId || !rewardDeductionFieldId || !cancelFeeFieldId || !urgentFeeFieldId) {
      throw new Error('源表中缺少必要的字段');
    }
    
    // 获取字段对象
    const productCategoryField = await sourceTable.getField<ISingleSelectField>(productCategoryFieldId);
    const outDateField = await sourceTable.getField<IDateTimeField>(outDateFieldId);
    const sharesField = await sourceTable.getField<INumberField>(sharesFieldId);
    const preparationFeeField = await sourceTable.getField<ICurrencyField>(preparationFeeFieldId);
    const rewardDeductionField = await sourceTable.getField<ICurrencyField>(rewardDeductionFieldId);
    const cancelFeeField = await sourceTable.getField<ICurrencyField>(cancelFeeFieldId);
    const urgentFeeField = await sourceTable.getField<ICurrencyField>(urgentFeeFieldId);
    
    // 获取记录ID列表
    const recordIdList = await sourceTable.getRecordIdList();
    
    // 分组数据
    const groupedData: Record<string, {
      count: number;
      shares: number;
      preparationFee: number;
      rewardDeduction: number;
      cancelFee: number;
      urgentFee: number;
    }> = {};
    
    // 遍历记录
    for (const recordId of recordIdList) {
      // 获取日期值
      const outDateValue = await outDateField.getValue(recordId);
      
      // 检查是否有日期值
      if (!outDateValue) continue;
      
      // 解析日期
      const outDate = new Date(outDateValue);
      const outDateMonth = outDate.getMonth() + 1; // 月份从0开始，需要+1
      
      // 筛选指定月份的记录
      if (outDateMonth !== month) continue;
      
      // 获取产品类别
      const productCategory = await productCategoryField.getValue(recordId);
      
      // 如果没有产品类别则跳过
      if (!productCategory) continue;
      
      // 初始化分组数据
      if (!groupedData[productCategory]) {
        groupedData[productCategory] = {
          count: 0,
          shares: 0,
          preparationFee: 0,
          rewardDeduction: 0,
          cancelFee: 0,
          urgentFee: 0
        };
      }
      
      // 累加计数
      groupedData[productCategory].count += 1;
      
      // 累加份数
      const sharesValue = await sharesField.getValue(recordId);
      if (sharesValue !== null && sharesValue !== undefined) {
        groupedData[productCategory].shares += Number(sharesValue);
      }
      
      // 累加制备费
      const preparationFeeValue = await preparationFeeField.getValue(recordId);
      if (preparationFeeValue !== null && preparationFeeValue !== undefined) {
        groupedData[productCategory].preparationFee += Number(preparationFeeValue);
      }
      
      // 累加出库订单奖励扣减
      const rewardDeductionValue = await rewardDeductionField.getValue(recordId);
      if (rewardDeductionValue !== null && rewardDeductionValue !== undefined) {
        groupedData[productCategory].rewardDeduction += Number(rewardDeductionValue);
      }
      
      // 累加制剂取消收取费用
      const cancelFeeValue = await cancelFeeField.getValue(recordId);
      if (cancelFeeValue !== null && cancelFeeValue !== undefined) {
        groupedData[productCategory].cancelFee += Number(cancelFeeValue);
      }
      
      // 累加制剂加急/非正常出库时间收取费用
      const urgentFeeValue = await urgentFeeField.getValue(recordId);
      if (urgentFeeValue !== null && urgentFeeValue !== undefined) {
        groupedData[productCategory].urgentFee += Number(urgentFeeValue);
      }
    }
    
    return groupedData;
  }
  
  /**
   * 将数据写入新表
   */
  private async writeDataToNewTable(newTable: ITable, groupedData: Record<string, any>): Promise<void> {
    // 获取字段元数据
    const fields = await newTable.getFieldMetaList();
    
    // 获取字段ID
    const projectFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.PROJECT)?.id;
    const personCountFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.PERSON_COUNT)?.id;
    const sharesFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.SHARES)?.id;
    const preparationFeeFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.PREPARATION_FEE)?.id;
    const rewardDeductionFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.REWARD_DEDUCTION)?.id;
    const cancelFeeFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.CANCEL_FEE)?.id;
    const urgentFeeFieldId = fields.find(field => field.name === this.TARGET_FIELD_NAMES.URGENT_FEE)?.id;
    
    // 验证字段是否存在
    if (!projectFieldId || !personCountFieldId || !sharesFieldId || 
        !preparationFeeFieldId || !rewardDeductionFieldId || !cancelFeeFieldId || !urgentFeeFieldId) {
      throw new Error('新表中缺少必要的字段');
    }
    
    // 准备记录数据
    const records = Object.entries(groupedData).map(([productCategory, data]) => {
      return {
        fields: {
          [projectFieldId]: productCategory,
          [personCountFieldId]: data.count,
          [sharesFieldId]: data.shares.toFixed(2),
          [preparationFeeFieldId]: data.preparationFee.toFixed(2),
          [rewardDeductionFieldId]: data.rewardDeduction.toFixed(2),
          [cancelFeeFieldId]: data.cancelFee.toFixed(2),
          [urgentFeeFieldId]: data.urgentFee.toFixed(2)
        }
      };
    });
    
    // 批量添加记录
    if (records.length > 0) {
      await newTable.addRecords(records);
    }
  }
} 