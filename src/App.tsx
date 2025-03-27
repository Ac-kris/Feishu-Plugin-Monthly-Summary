import { useEffect, useState } from 'react';
import { bitable, ITable } from '@lark-base-open/js-sdk';
import { Select, Button, message, Spin } from 'antd';
import { MonthlyReportGenerator } from './MonthlyReportGenerator';
import './App.css';

interface TableOption {
  label: string;
  value: string;
}

interface MonthOption {
  label: string;
  value: number;
}

// 月份选项
const MONTH_OPTIONS: MonthOption[] = Array.from({ length: 12 }, (_, index) => ({
  label: `${index + 1}月`,
  value: index + 1,
}));

function App() {
  const [tables, setTables] = useState<TableOption[]>([]);
  const [selectedTableId, setSelectedTableId] = useState<string>('');
  const [selectedMonth, setSelectedMonth] = useState<number>(new Date().getMonth() + 1);
  const [loading, setLoading] = useState<boolean>(false);
  const [messageApi, contextHolder] = message.useMessage();

  // 获取表格列表
  useEffect(() => {
    async function loadTables() {
      try {
        const tableList = await bitable.base.getTables();
        const tableOptions = tableList.map((table) => ({
          label: table.name,
          value: table.id
        }));
        
        setTables(tableOptions);
        
        if (tableOptions.length > 0) {
          setSelectedTableId(tableOptions[0].value);
        }
      } catch (error) {
        messageApi.error('获取表格列表失败');
        console.error('获取表格列表失败:', error);
      }
    }
    
    loadTables();
  }, [messageApi]);

  // 生成月度汇总报告
  const generateReport = async () => {
    if (!selectedTableId) {
      messageApi.warning('请选择原数据表');
      return;
    }

    if (!selectedMonth) {
      messageApi.warning('请选择月份');
      return;
    }

    setLoading(true);
    
    try {
      const generator = new MonthlyReportGenerator();
      await generator.generate(selectedTableId, selectedMonth);
      messageApi.success('汇总表生成成功！');
    } catch (error) {
      messageApi.error('生成汇总表失败，请查看控制台');
      console.error('生成汇总表失败:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="container">
      {contextHolder}
      
      <div className="form-item">
        <label>选择原数据表</label>
        <Select 
          style={{ width: '100%' }}
          options={tables}
          value={selectedTableId}
          onChange={setSelectedTableId}
          placeholder="请选择原数据表"
        />
      </div>
      
      <div className="form-item">
        <label>选择月份</label>
        <Select 
          style={{ width: '100%' }}
          options={MONTH_OPTIONS}
          value={selectedMonth}
          onChange={setSelectedMonth}
          placeholder="请选择月份"
        />
      </div>
      
      <Button 
        type="primary" 
        onClick={generateReport}
        loading={loading}
        style={{ width: '100%' }}
      >
        生成
      </Button>
      
      {loading && (
        <div className="loading-container">
          <Spin tip="生成中..." />
        </div>
      )}
    </div>
  );
}

export default App; 