import { useEffect, useState } from 'react';
import { bitable } from '@lark-base-open/js-sdk';
import { Select, Button, message, Spin, Alert } from 'antd';
import { MonthlyReportGenerator } from './MonthlyReportGenerator';
import './App.css';

interface TableOption {
  label: string;
  value: string;
}

// 定义表格信息接口，与API返回值匹配
interface TableMeta {
  id: string;
  name: string;
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
  const [initialLoading, setInitialLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [messageApi, contextHolder] = message.useMessage();

  // 获取表格列表
  useEffect(() => {
    async function loadTables() {
      setInitialLoading(true);
      setErrorMessage(null);
      
      try {
        // 获取表格列表元数据
        const tableMetaList = await bitable.base.getTableMetaList();
        
        if (!tableMetaList || tableMetaList.length === 0) {
          setErrorMessage('未找到任何数据表，请先创建数据表');
          setInitialLoading(false);
          return;
        }
        
        const tableOptions = tableMetaList.map((table: TableMeta) => ({
          label: table.name || '未命名表格',
          value: table.id
        }));
        
        setTables(tableOptions);
        
        if (tableOptions.length > 0) {
          setSelectedTableId(tableOptions[0].value);
        }
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : '未知错误';
        setErrorMessage(`获取表格列表失败: ${errorMsg}`);
        messageApi.error('获取表格列表失败，请检查网络连接或刷新页面重试');
        console.error('获取表格列表失败:', error);
      } finally {
        setInitialLoading(false);
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
    setErrorMessage(null);
    
    try {
      const generator = new MonthlyReportGenerator();
      await generator.generate(selectedTableId, selectedMonth);
      messageApi.success('汇总表生成成功！');
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : '未知错误';
      setErrorMessage(`生成汇总表失败: ${errorMsg}`);
      messageApi.error('生成汇总表失败，请查看详细错误信息');
      console.error('生成汇总表失败:', error);
    } finally {
      setLoading(false);
    }
  };

  if (initialLoading) {
    return (
      <div className="container loading-container">
        <Spin tip="加载中..." size="large" />
      </div>
    );
  }

  return (
    <div className="container">
      {contextHolder}
      
      {errorMessage && (
        <div className="error-message">
          <Alert
            message="错误"
            description={errorMessage}
            type="error"
            showIcon
            closable
            onClose={() => setErrorMessage(null)}
          />
        </div>
      )}
      
      <div className="form-item">
        <label>选择原数据表</label>
        <Select 
          style={{ width: '100%' }}
          options={tables}
          value={selectedTableId}
          onChange={setSelectedTableId}
          placeholder="请选择原数据表"
          disabled={loading || tables.length === 0}
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
          disabled={loading}
        />
      </div>
      
      <Button 
        type="primary" 
        onClick={generateReport}
        loading={loading}
        style={{ width: '100%' }}
        disabled={!selectedTableId || loading}
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