
SELECT column1, column2 
FROM test_table;






SELECT * FROM another_table;


SELECT 
  column1, 
  
 
  column2,
  column3 
FROM third_table
WHERE column1 = '�e�L�X�g'; 

CREATE OR REPLACE PROCEDURE test_proc
IS
  v_variable VARCHAR2(100) := '����� /* �R�����g�ł͂���܂��� */ ������ł�'; 
  v_another  VARCHAR2(100) := '����� -- �R�����g�ł͂���܂���';
BEGIN
  


 SELECT * FROM table1;
  
  
  FOR i IN 1..10 LOOP
    dbms_output.put_line('�J�E���g: ' || i); 
  END LOOP;
  
   SELECT 1 FROM dual; 
  
  IF v_variable = '�e�X�g' THEN
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
  '�e�L�X�g/* ����̓R�����g�ł͂Ȃ� */' AS column1,
  '�e�L�X�g-- ������R�����g�ł͂Ȃ�' AS column2
FROM dual;







BEGIN
  NULL; 
END;
/

