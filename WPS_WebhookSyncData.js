// 测试数据生成函数
function generateTestData() {
    console.log("[INFO] 生成测试数据...");
    return {
          records: [
            {
              '工单key': 'LIN-5', 
              '工单类型': '问题记录', 
              '工单概要': '小林忘记设置闹钟了', 
              '客户名称': '小林之家', 
              '创建日期': '2024-12-31', 
              '优先级': '高', 
              '报告人': '小林他妈', 
              '经办人': '小林', 
              '状态': '已解决', 
              '解决结果': '已给出解决方案',
              '已解决时间': '2024-12-31'
            }, 
            {
              '工单key': 'LIN-6', 
              '工单类型': '问题记录', 
              '工单概要': '小林上学迟到了', 
              '客户名称': '小林之家', 
              '创建日期': '2024-12-31', 
              '优先级': '高', 
              '报告人': '小林老师', 
              '经办人': '小林他妈', 
              '状态': '已解决', 
              '解决结果': '已给出解决方案',
              '已解决时间': '2024-12-31'
            }
          ]
    };
}

try {
    // 测试数据开关（true: 使用测试数据 false: 使用真实数据）
    const USE_TEST_DATA = true ; // 默认关闭测试数据
    
    // ================== 动态获取表ID ==================
    const TARGET_SHEET_NAME = "小林动态报告"; // 需要同步的目标表名称
    const sheets = Application.Sheet.GetSheets();
    let CURRENT_SHEET_ID = null; 

    // 遍历所有表，查找目标表
    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        if (sheet.name === TARGET_SHEET_NAME) {
            CURRENT_SHEET_ID = sheet.id;
            console.log(`[INFO] 获取表ID成功: ${CURRENT_SHEET_ID}`);
            break; // 找到后退出循环
        }
    }
    // 添加表存在性检查
    if (!CURRENT_SHEET_ID) {
        throw new Error(`[ERROR] 未找到名称为 ${TARGET_SHEET_NAME} 的表`);
    }
       
    // ================== Webhook数据检查 ==================
    let jiraData;
    if (Context.argv && Context.argv.records) { // 检查Webhook数据是否存在
        try {
            jiraData = Context.argv;
            console.log("[INFO] 接收到Webhook数据:", JSON.stringify(jiraData, null, 2));
        } catch (e) {
            console.error("[ERROR] Webhook数据解析失败:", e.message);
            throw new Error("无效的JSON格式数据");
        }
    } else {
        console.warn("[WARN] 未接收到有效Webhook数据，启用备用数据源");
        jiraData = USE_TEST_DATA ? generateTestData() : { records: [] };
    }
    
    // ================== 获取全量现有记录（分页逻辑） ==================
    let existingRecords = [];
    let offset = null;
    const PAGE_SIZE = 1000; // 使用API允许的最大分页值

    try {
    while (true) {
        const result = Application.Record.GetRecords({
            SheetId: CURRENT_SHEET_ID,
            Offset: offset,
            PageSize: PAGE_SIZE // 明确指定分页大小
        });

        if (!result.records || result.records.length === 0) break;
        
        existingRecords = existingRecords.concat(result.records);
        console.log(`[DEBUG] 已加载 ${existingRecords.length} 条记录`);

        // 通过offset判断是否继续请求
        offset = result.offset;
        if (!offset) break;
    }
    console.log(`[SUCCESS] 共获取 ${existingRecords.length} 条现有记录`);
    } catch (e) {
      throw new Error(`获取现有记录失败: ${e.message}`);
    }

    // ================== 主业务逻辑 ==================
    // const existingRecords = Application.Record.GetRecords({ SheetId: CURRENT_SHEET_ID }).records;
    // console.log(existingRecords.length)
    // 构建工单KEY映射表（用于快速查找）
    const existingKeyMap = new Map();
    existingRecords.forEach(record => {
        const key = record.fields["工单key"];
        existingKeyMap.set(key, record);
    });
    
    const jiraRecords = jiraData.records || [];
    // 处理新增/更新逻辑
    const recordsToUpdate = [];
    const recordsToCreate = [];

    jiraRecords.forEach(jiraRecord => {
        const existingRecord = existingKeyMap.get(jiraRecord["工单key"]);
        if (existingRecord) {
            // 检查字段是否更新
            let needUpdate = false;
            const updateFields = {};
            Object.keys(jiraRecord).forEach(key => {
                const jiraValue = jiraRecord[key];
                const existingValue = existingRecord.fields[key];
                if (JSON.stringify(jiraValue) !== JSON.stringify(existingValue)) {
                    needUpdate = true;
                    updateFields[key] = jiraValue;
                }
            });
            if (needUpdate) {
                recordsToUpdate.push({
                    id: existingRecord.id,
                    fields: updateFields
                });
            }
        } else {
            // 新增记录
            recordsToCreate.push({
                fields: jiraRecord
            });
        }
    });

    // 执行更新操作
    if (recordsToUpdate.length > 0) {
        console.log(`[INFO] 更新 ${recordsToUpdate.length} 条记录`);
        Application.Record.UpdateRecords({
            SheetId: CURRENT_SHEET_ID,
            Records: recordsToUpdate
        });
    }

    // 执行新增操作
    if (recordsToCreate.length > 0) {
        console.log(`[INFO] 新增 ${recordsToCreate.length} 条记录`);
        Application.Record.CreateRecords({
            SheetId: CURRENT_SHEET_ID,
            Records: recordsToCreate
        });
    }

    console.log("[SUCCESS] 数据同步完成");
} catch (error) {
    console.error("[ERROR] 同步失败:", error.message);
    throw error; // 抛出错误以便在WPS日志中查看
}
