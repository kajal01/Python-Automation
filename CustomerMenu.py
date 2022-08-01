from subprocess import call
import os.path

print("Running Customer Menu Test Cases...")

#execute python code

my_path = os.path.abspath(os.path.dirname(__file__))

call(["python", os.path.join(my_path, "Customer\AddCustomer.py")])

