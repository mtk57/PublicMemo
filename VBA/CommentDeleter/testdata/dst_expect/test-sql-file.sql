
SELECT column1, column2 
FROM test_table;






SELECT * FROM another_table;


SELECT 
  column1, 
  
 
  column2,
  column3 
FROM third_table
WHERE column1 = 'テキスト'; 

CREATE OR REPLACE PROCEDURE test_proc
IS
  v_variable VARCHAR2(100) := 'これは /* コメントではありません */ 文字列です'; 
  v_another  VARCHAR2(100) := 'これも -- コメントではありません';
BEGIN
  


 SELECT * FROM table1;
  
  
  FOR i IN 1..10 LOOP
    dbms_output.put_line('カウント: ' || i); 
  END LOOP;
  
   SELECT 1 FROM dual; 
  
  IF v_variable = 'テスト' THEN
    NULL; 
  END IF;
END;
/

CREATE OR REPLACE FUNCTION test_func(p_param1 IN NUMBER) 
RETURN NUMBER
IS
  
  v_result NUMBER;
BEGIN
  v_result := p_param1 * 2; 
  
  





  
  RETURN v_result;
END;
/

SELECT 
  'テキスト/* これはコメントではない */' AS column1,
  'テキスト-- これもコメントではない' AS column2
FROM dual;







BEGIN
  NULL; 
END;
/

