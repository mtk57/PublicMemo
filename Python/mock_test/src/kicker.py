#!/usr/bin/env python3


from child import Child


ins_c = Child()

print(f'Child ins mem={ins_c.child_func()},  child class mem={Child.child_class_func()}')