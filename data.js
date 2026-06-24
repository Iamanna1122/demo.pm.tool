window.INITIAL_DATA = {
  "users": [
    { "username": "admin", "password": "123", "role": "Admin", "name": "系統管理員", "dept": "資訊部" },
    { "username": "pm_user", "password": "123", "role": "PM", "name": "曼玉", "dept": "產品管理部" },
    { "username": "rd_user", "password": "123", "role": "RD", "name": "世偉", "dept": "研發部" },
    { "username": "rd_user2", "password": "123", "role": "RD", "name": "齊政", "dept": "研發部" },
    { "username": "vp_user", "password": "123", "role": "VP", "name": "副總經理", "dept": "總經辦" }
  ],
  "projects": [
    {
      "id": "SIH-145-CONV",
      "name": "14.5 常規",
      "type": "Conventional",
      "currentVersion": "V9.1（量產）／V8（試戴）",
      "priority": 1,
      "gate": "G5",
      "status": "黃",
      "bottleneck": "仍需確認無印痕、漏白；V8 試戴討論日期有兩種說法",
      "nextStep": "確認 V9.1 生產良率與缺陷；完成 V8 全光度試戴結論",
      "owner": "曼玉",
      "collaborators": "世偉、齊政、鍾民",
      "deadline": "2026-06-24",
      "statusDesc": "V9.1 已安排生產",
      "lastUpdate": "2026-06-23",
      "confidence": "中",
      "note": "需確認 V8 與 V9.1 的用途及版本關係",
      "milestones": [
        { "name": "G0 規格凍結", "target": "2025-11-03", "actual": "2025-11-03", "status": "green" },
        { "name": "G1 平光設計", "target": "2025-11-17", "actual": "2025-11-17", "status": "green" },
        { "name": "G2 平光測試", "target": "2025-12-15", "actual": "2025-12-15", "status": "green" },
        { "name": "G3 全光設計", "target": "2026-02-02", "actual": "2026-02-02", "status": "green" },
        { "name": "G4 全光試戴", "target": "2026-05-11", "actual": "2026-05-25", "status": "green" },
        { "name": "G5 試量產", "target": "2026-06-22", "actual": "2026-06-22", "status": "blue" },
        { "name": "G6 正式量產", "target": "2026-06-29", "actual": "", "status": "gray" }
      ]
    },
    {
      "id": "SIH-145-SIH",
      "name": "14.5 細水膠",
      "type": "SiH",
      "currentVersion": "V8",
      "priority": 2,
      "gate": "G3",
      "status": "黃",
      "bottleneck": "V7 油墨誤用掉色；全光度模仁需 3–4 天；氣孔數據來源不一致",
      "nextStep": "重做彩片、確認油墨配方、完成全光度模仁並排入試戴",
      "owner": "曼玉",
      "collaborators": "世偉、鍾民、齊政",
      "deadline": "2026-06-25",
      "statusDesc": "尚未量產",
      "lastUpdate": "2026-06-23",
      "confidence": "中",
      "note": "下一階段優先開發產品",
      "milestones": [
        { "name": "G0 規格凍結", "target": "2025-11-10", "actual": "2025-11-10", "status": "green" },
        { "name": "G1 平光設計", "target": "2025-11-24", "actual": "2025-11-24", "status": "green" },
        { "name": "G2 平光測試", "target": "2025-12-22", "actual": "2025-12-22", "status": "green" },
        { "name": "G3 全光設計", "target": "2026-06-23", "actual": "2026-06-23", "status": "blue" },
        { "name": "G4 全光試戴", "target": "2026-07-15", "actual": "", "status": "gray" },
        { "name": "G5 試量產", "target": "2026-08-10", "actual": "", "status": "gray" },
        { "name": "G6 正式量產", "target": "2026-08-30", "actual": "", "status": "gray" }
      ]
    },
    {
      "id": "SIH-142-SIH",
      "name": "14.2 細水膠",
      "type": "SiH",
      "currentVersion": "V18.2",
      "priority": 3,
      "gate": "G4",
      "status": "黃",
      "bottleneck": "樣本僅 8–12 人；需建立 1–8 與 9–16 區域差異的判讀規則",
      "nextStep": "建立全光度試戴計畫並與視陽基準比較",
      "owner": "曼玉",
      "collaborators": "齊政、宣義、鍾民",
      "deadline": "2026-06-25",
      "statusDesc": "尚未量產",
      "lastUpdate": "2026-06-23",
      "confidence": "中",
      "note": "14.5 系列後推進；超薄優先於常規",
      "milestones": [
        { "name": "G0 規格凍結", "target": "2025-11-03", "actual": "2025-11-03", "status": "green" },
        { "name": "G1 平光設計", "target": "2025-11-17", "actual": "2025-11-17", "status": "green" },
        { "name": "G2 平光測試", "target": "2025-12-15", "actual": "2025-12-15", "status": "green" },
        { "name": "G3 全光設計", "target": "2026-02-09", "actual": "2026-02-09", "status": "green" },
        { "name": "G4 全光試戴", "target": "2026-06-23", "actual": "2026-06-23", "status": "blue" },
        { "name": "G5 試量產", "target": "2026-07-20", "actual": "", "status": "gray" },
        { "name": "G6 正式量產", "target": "2026-08-15", "actual": "", "status": "gray" }
      ]
    },
    {
      "id": "SIH-142-UT",
      "name": "14.2 超薄",
      "type": "SiH",
      "currentVersion": "V14",
      "priority": 4,
      "gate": "G4",
      "status": "紅",
      "bottleneck": "同名 V14 另被記載為嚴重漏白、不可量產；版本身分不清",
      "nextStep": "先確認版本代號，再執行 6/26 試戴與底部凸起評估",
      "owner": "曼玉",
      "collaborators": "齊政、鍾民",
      "deadline": "2026-06-25",
      "statusDesc": "量產適用性待確認",
      "lastUpdate": "2026-06-23",
      "confidence": "低",
      "note": "版本命名衝突，必須拆清產品／模仁組合",
      "milestones": [
        { "name": "G0 規格凍結", "target": "2025-11-03", "actual": "2025-11-03", "status": "green" },
        { "name": "G1 平光設計", "target": "2025-11-17", "actual": "2025-11-17", "status": "green" },
        { "name": "G2 平光測試", "target": "2025-12-15", "actual": "2025-12-15", "status": "green" },
        { "name": "G3 全光設計", "target": "2026-02-09", "actual": "2026-02-09", "status": "green" },
        { "name": "G4 全光試戴", "target": "2026-06-26", "actual": "", "status": "red" },
        { "name": "G5 試量產", "target": "2026-07-25", "actual": "", "status": "gray" },
        { "name": "G6 正式量產", "target": "2026-08-20", "actual": "", "status": "gray" }
      ]
    },
    {
      "id": "SIH-142-CONV",
      "name": "14.2 常規",
      "type": "Conventional",
      "currentVersion": "V4",
      "priority": 5,
      "gate": "G4",
      "status": "黃",
      "bottleneck": "紀錄稱全光度尚未車製，但又排定 6/26 試戴",
      "nextStep": "確認試戴樣片類型並評估底部凸起、硬痕與漏白",
      "owner": "曼玉",
      "collaborators": "齊政、鍾民",
      "deadline": "2026-06-25",
      "statusDesc": "既有量產版本仍有品質疑慮",
      "lastUpdate": "2026-06-23",
      "confidence": "低",
      "note": "需確認「常規」是否為矽水膠或傳統水膠",
      "milestones": [
        { "name": "G0 規格凍結", "target": "2025-11-03", "actual": "2025-11-03", "status": "green" },
        { "name": "G1 平光設計", "target": "2025-11-17", "actual": "2025-11-17", "status": "green" },
        { "name": "G2 平光測試", "target": "2025-12-15", "actual": "2025-12-15", "status": "green" },
        { "name": "G3 全光設計", "target": "2026-02-09", "actual": "2026-02-09", "status": "green" },
        { "name": "G4 全光試戴", "target": "2026-06-26", "actual": "", "status": "blue" },
        { "name": "G5 試量產", "target": "2026-07-20", "actual": "", "status": "gray" },
        { "name": "G6 正式量產", "target": "2026-08-15", "actual": "", "status": "gray" }
      ]
    }
  ],
  "versions": [
    { "projectId": "SIH-145-CONV", "version": "V9.1", "category": "量產版", "change": "改善氣孔；邊緣切角有變化", "testType": "生產／切片", "sampleSize": "", "porosity": "", "comfort": "", "defect": "目標無印痕、漏白", "conclusion": "生產中，待良率與缺陷結果", "nextStep": "回收首批生產數據", "owner": "世偉", "date": "2026-06-22", "source": "廠內製程／會議紀錄", "isProduction": "條件式可用", "note": "量產版本" },
    { "projectId": "SIH-145-CONV", "version": "V8", "category": "試戴版", "change": "會議紀錄未明確說明", "testType": "全光度試戴", "sampleSize": "", "porosity": "", "comfort": "", "defect": "待確認", "conclusion": "待決策", "nextStep": "確認 6/23 或 6/25 的正式評估結論", "owner": "齊政", "date": "2026-06-25", "source": "內部試戴", "isProduction": "待確認", "note": "日期有衝突" },
    { "projectId": "SIH-145-SIH", "version": "V7", "category": "平光彩片", "change": "原構型", "testType": "平光／製程", "sampleSize": "", "porosity": "0.75", "comfort": "", "defect": "油墨誤用、氣孔", "conclusion": "NG", "nextStep": "重做彩片並改用正確油墨", "owner": "世偉", "date": "", "source": "製程紀錄", "isProduction": "否", "note": "氣孔率 70–80%，以 75% 暫列中點" },
    { "projectId": "SIH-145-SIH", "version": "V8", "category": "平光樣片", "change": "構型未改；調整邊緣以利排氣", "testType": "平光／製程", "sampleSize": "", "porosity": "0.05", "comfort": "", "defect": "氣孔仍需持續確認", "conclusion": "可進全光度", "nextStep": "車製全光度模仁", "owner": "鍾民", "date": "2026-06-23", "source": "製程測試", "isProduction": "否", "note": "製程測 5.0%；另有世偉測 1.3%" },
    { "projectId": "SIH-142-SIH", "version": "V18.2", "category": "候選版", "change": "改善硬痕、漏白及底部凸起", "testType": "臨床試戴", "sampleSize": "10", "porosity": "0.045", "comfort": "4.2", "defect": "樣本數偏少", "conclusion": "可進全光度試戴", "nextStep": "擴大樣本並與競品比較", "owner": "齊政", "date": "2026-06-23", "source": "臨床世代結果", "isProduction": "待確認", "note": "需保留區域分組定義" },
    { "projectId": "SIH-142-SIH", "version": "V14", "category": "前版", "change": "處理硬痕／漏白", "testType": "試戴／製程", "sampleSize": "", "porosity": "", "comfort": "", "defect": "嚴重漏白", "conclusion": "NG", "nextStep": "確認是否與超薄 V14 為同一代號", "owner": "鍾民", "date": "", "source": "會議紀錄", "isProduction": "否", "note": "版本代號可能衝突" },
    { "projectId": "SIH-142-UT", "version": "V14", "category": "候選版／待辨識", "change": "優化底部凸起", "testType": "試戴", "sampleSize": "", "porosity": "", "comfort": "", "defect": "另有 V14 嚴重漏白紀錄", "conclusion": "阻塞待確認", "nextStep": "先完成版本辨識", "owner": "曼玉", "date": "2026-06-26", "source": "內部試戴排程", "isProduction": "否", "note": "不得在代號未釐清前判定量產" },
    { "projectId": "SIH-142-CONV", "version": "V4", "category": "候選版", "change": "優化底部凸起", "testType": "試戴", "sampleSize": "", "porosity": "", "comfort": "", "defect": "硬痕、漏白、底部凸起", "conclusion": "待測", "nextStep": "確認樣片是否為全光度", "owner": "齊政", "date": "2026-06-26", "source": "內部試戴排程", "isProduction": "待確認", "note": "全光度車製狀態矛盾" }
  ],
  "actions": [
    { "id": "A-001", "projectId": "SIH-145-SIH", "title": "油墨重做與防呆", "desc": "針對 V7 掉色問題，重做彩片並建立配方領用防呆", "owner": "世偉", "coowners": "品保、生產", "priority": "高", "deadline": "2026-06-26", "status": "進行中", "notes": "攸關臨床安全性" },
    { "id": "A-002", "projectId": "SIH-145-SIH", "title": "對齊氣孔率標準", "desc": "解決 5.0% 與 1.3% 的測量定義衝突，召開會議討論統一抽樣", "owner": "宣義", "coowners": "世偉、品質部", "priority": "高", "deadline": "2026-06-26", "status": "進行中", "notes": "避免 Gate 判定失真" },
    { "id": "A-003", "projectId": "SIH-142-SIH", "title": "建立 14.2 SiH 判讀規則", "desc": "針對 1-8 穴與 9-16 穴區域差異，建立量產缺陷判定標準", "owner": "齊政", "coowners": "世偉、品保", "priority": "高", "deadline": "2026-06-26", "status": "進行中", "notes": "試戴分析必要前置" },
    { "id": "A-004", "projectId": "SIH-142-SIH", "title": "全光度試戴計畫規劃", "desc": "撰寫 V18.2 全光度臨床計畫書，明訂樣本數與比較基準", "owner": "齊政", "coowners": "宣義、設計組", "priority": "高", "deadline": "2026-06-26", "status": "進行中", "notes": "下週審查準備" },
    { "id": "A-009", "projectId": "SIH-142-UT", "title": "釐清 V14 身分與可用性", "desc": "確認是否與 14.2 細水膠嚴重漏白 V14 為同一版本", "owner": "曼玉", "coowners": "鍾民、齊政", "priority": "高", "deadline": "2026-06-24", "status": "未開始", "notes": "未釐清前不可做量產判定" }
  ],
  "risks": [
    { "id": "RI-001", "type": "Issue", "projectId": "SIH-145-SIH", "title": "V7 使用水膠油墨造成掉色", "desc": "彩片需重做，有品質與客戶安全風險", "probability": 5, "impact": 5, "score": 25, "light": "紅", "mitigation": "立即修正油墨、重做並建立領料／配方防呆", "owner": "世偉", "deadline": "2026-06-26", "status": "處理中" },
    { "id": "RI-002", "type": "Risk", "projectId": "SIH-145-SIH", "title": "氣孔率數據量測不一致", "desc": "5.0% 與 1.3% 不可直接比較，影響 Gate 判定", "probability": 4, "impact": 4, "score": 16, "light": "黃", "mitigation": "統一定義、抽樣與分母後再作 Gate 判定", "owner": "宣義", "deadline": "2026-06-26", "status": "待確認" },
    { "id": "RI-003", "type": "Issue", "projectId": "SIH-142-UT", "title": "V14 版本代號重複且結論矛盾", "desc": "一處為嚴重漏白 NG；一處排定試戴，影響量產判定", "probability": 5, "impact": 5, "score": 25, "light": "紅", "mitigation": "建立唯一版本碼並由設計端確認版本 lineage", "owner": "曼玉", "deadline": "2026-06-24", "status": "待確認" }
  ]
};
