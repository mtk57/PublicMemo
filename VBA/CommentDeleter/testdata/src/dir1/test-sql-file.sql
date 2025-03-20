-- これは単一行コメントのテストです
SELECT column1, column2 -- 行末コメント
FROM test_table;

/* これは複数行コメントの
   テストです。複数行に
   わたるコメントを
   正しく削除できるか確認します。
*/
SELECT * FROM another_table;

-- コメント内に引用符 'テスト' がある場合
SELECT 
  column1, 
  /* 行の途中からの
     コメント */ 
  column2,
  column3 -- もう一つの行末コメント
FROM third_table
WHERE column1 = 'テキスト'; -- 文字列の後のコメント

CREATE OR REPLACE PROCEDURE test_proc
IS
  v_variable VARCHAR2(100) := 'これは /* コメントではありません */ 文字列です'; -- 文字列内のコメント記号
  v_another  VARCHAR2(100) := 'これも -- コメントではありません';
BEGIN
  /* コメント開始
  複数の
  行にわたる
  コメント */ SELECT * FROM table1;
  
  -- プロシージャの処理内容
  FOR i IN 1..10 LOOP
    dbms_output.put_line('カウント: ' || i); -- ループ内コメント
  END LOOP;
  
  /* 複数行コメント1 */ SELECT 1 FROM dual; /* 複数行コメント2 */
  
  IF v_variable = 'テスト' THEN
    NULL; -- 空の処理
  END IF;
END;
/

CREATE OR REPLACE FUNCTION test_func(p_param1 IN NUMBER) -- パラメータの説明コメント
RETURN NUMBER
IS
  /* 変数宣言部分のコメント */
  v_result NUMBER;
BEGIN
  v_result := p_param1 * 2; -- 2倍にする
  
  /* 
  複雑な
  -- ネストしたコメント（行コメントを含む複数行コメント）
  計算の
  説明 
  */
  
  RETURN v_result;
END;
/

SELECT 
  'テキスト/* これはコメントではない */' AS column1,
  'テキスト-- これもコメントではない' AS column2
FROM dual;

-- 行コメントの後の空行

/* 
複数行コメントの後の空行
*/

BEGIN
  NULL; -- コメント
END;
/
