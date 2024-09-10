import subprocess

def run_tests():
    # List of test files to run
    test_files = [
        "tests/test_pdf_meta.py",
        "tests/test_pptx_meta.py",
        "tests/test_docx_meta.py",
    ]
    
    # File to store full test output
    full_output_file = "test_full_output.txt"
    
    # File to store test failure summary
    failure_summary_file = "test_failure_summary.txt"
    
    # Run the tests and capture the output
    with open(full_output_file, "w") as full_output, open(failure_summary_file, "w") as failure_output:
        all_tests_passed = True
        
        for test_file in test_files:
            print(f"Running {test_file}...")
            # Run each test and capture the output
            result = subprocess.run(
                ["pytest", test_file, "-vvv"], 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            
            # Write the full output to the log file
            full_output.write(f"===== Running {test_file} =====\n")
            full_output.write(result.stdout)
            full_output.write(result.stderr)
            full_output.write("\n\n")

            # Check if there were any failures
            if "FAILURES" in result.stdout or "FAILURES" in result.stderr:
                all_tests_passed = False
                # Write the failure summary to the second file
                failure_output.write(f"===== Failures in {test_file} =====\n")
                failure_output.write(result.stdout)
                failure_output.write(result.stderr)
                failure_output.write("\n\n")
        
        # If no failures, write "No failures" to the failure log
        if all_tests_passed:
            failure_output.write("No failures\n")
    
    print(f"Full output saved to {full_output_file}")
    print(f"Failure summary saved to {failure_summary_file}")

if __name__ == "__main__":
    run_tests()
