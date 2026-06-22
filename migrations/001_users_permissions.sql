-- ユーザー管理テーブル
CREATE TABLE IF NOT EXISTS users (
    id            uuid        DEFAULT gen_random_uuid() PRIMARY KEY,
    username      text        UNIQUE NOT NULL,
    display_name  text        NOT NULL DEFAULT '',
    password_hash text        NOT NULL,
    role          text        NOT NULL DEFAULT 'viewer'
                              CHECK (role IN ('admin', 'editor', 'viewer')),
    created_at    timestamptz DEFAULT now(),
    updated_at    timestamptz DEFAULT now()
);

-- プロジェクト閲覧権限テーブル
CREATE TABLE IF NOT EXISTS project_permissions (
    id          uuid        DEFAULT gen_random_uuid() PRIMARY KEY,
    user_id     uuid        NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    project_id  uuid        NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    created_at  timestamptz DEFAULT now(),
    UNIQUE(user_id, project_id)
);

-- インデックス
CREATE INDEX IF NOT EXISTS idx_project_permissions_user
    ON project_permissions(user_id);
CREATE INDEX IF NOT EXISTS idx_project_permissions_project
    ON project_permissions(project_id);
