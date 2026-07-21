"""TestAll 資料庫存取層 - SQLite 實作與 Protocol 契約

對外暴露：
- OCRReviewRepository (Protocol): 所有 Core Service 依賴的抽象介面
- SQLiteReviewRepository: SQLite 實作
- get_repo(): 快速取得 Repository 實例
- init_db(): 根據 schema.sql 初始化資料庫
"""