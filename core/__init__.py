"""TestAll 核心模組 - OCR 稽核服務, 庫存計算, 規則引擎

模組架構：
- ocr_review_service: 回存 Word + 入庫核心流程
- inventory_service: 每日庫存結算, 損耗判斷, 一致性驗證
- rule_engine: 內建/使用者自訂規則與警示管理
- db.repository: SQLite 資料存取層 (符合 OCRReviewRepository Protocol)
"""