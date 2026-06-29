import ast
import os
import sys

class SecurityAuditNodeVisitor(ast.NodeVisitor):
    def __init__(self, filename):
        self.filename = filename
        self.violations = []

    def report(self, node, message):
        self.violations.append(f"[{self.filename}:{node.lineno}] {message}")

    def visit_Call(self, node):
        # Check for direct calls to dangerous builtins
        if isinstance(node.func, ast.Name):
            if node.func.id in ['eval', 'exec']:
                self.report(node, f"DANGER: Use of {node.func.id}() detected. This can lead to Remote Code Execution.")
            elif node.func.id == 'open' and len(node.args) > 0:
                # Just a warning for open() if not carefully used, but usually fine. Skip to reduce noise.
                pass
                
        # Check for module attribute calls (e.g. os.system, subprocess.Popen)
        elif isinstance(node.func, ast.Attribute):
            if isinstance(node.func.value, ast.Name):
                module_name = node.func.value.id
                func_name = node.func.attr
                
                if module_name == 'os' and func_name in ['system', 'popen']:
                    self.report(node, f"DANGER: Use of os.{func_name}() detected. Potential command injection.")
                elif module_name == 'subprocess' and func_name in ['Popen', 'run', 'call', 'check_call', 'check_output']:
                    self.report(node, f"WARNING: Use of subprocess.{func_name}() detected. Audit inputs carefully.")
                elif module_name == 'pickle' and func_name in ['loads', 'load']:
                    self.report(node, f"DANGER: Use of pickle.{func_name}() detected. Insecure deserialization risk.")
                elif module_name == 'ast' and func_name == 'literal_eval':
                    self.report(node, "WARNING: ast.literal_eval() detected. Safer than eval(), but ensure inputs are trusted.")

        self.generic_visit(node)

    def visit_While(self, node):
        # Look for while True or equivalent
        is_infinite = False
        if isinstance(node.test, ast.Constant) and node.test.value is True:
            is_infinite = True
        elif isinstance(node.test, ast.NameConstant) and node.test.value is True: # Older Python ast support
            is_infinite = True
        elif isinstance(node.test, ast.Num) and node.test.n == 1:
            is_infinite = True
            
        if is_infinite:
            # Look for a break statement inside the while loop body
            has_break = any(isinstance(child, ast.Break) for child in ast.walk(node))
            if not has_break:
                self.report(node, "WARNING: Infinite 'while True' loop without a break detected. This will fatally freeze NVDA.")

        self.generic_visit(node)

def scan_directory(directory):
    total_violations = []
    
    for root, dirs, files in os.walk(directory):
        if '__pycache__' in root:
            continue
            
        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        source = f.read()
                    
                    tree = ast.parse(source, filename=filepath)
                    visitor = SecurityAuditNodeVisitor(filepath)
                    visitor.visit(tree)
                    
                    total_violations.extend(visitor.violations)
                except SyntaxError as e:
                    total_violations.append(f"[{filepath}:{e.lineno}] SYNTAX ERROR: Could not parse file.")
                except Exception as e:
                    total_violations.append(f"[{filepath}] ERROR: {str(e)}")
                    
    return total_violations

if __name__ == "__main__":
    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..', '..'))
    addon_dir = os.path.join(repo_root, 'addon')
    
    if not os.path.exists(addon_dir):
        print(f"Error: Could not find addon directory at {addon_dir}")
        sys.exit(1)
        
    print("==================================================")
    print("BOA SECURITY AUDITOR (AST-BASED SCANNER)")
    print("==================================================")
    print(f"Scanning target: {addon_dir}\n")
    
    violations = scan_directory(addon_dir)
    
    if not violations:
        print("[PASS] AUDIT PASSED: No security vulnerabilities or unsafe patterns detected.")
    else:
        print("[WARN] AUDIT WARNING: The following potential vulnerabilities were flagged:\n")
        for v in violations:
            print(v)
        print("\nNote: Please review these flags manually. Do not attempt to fix them automatically.")
