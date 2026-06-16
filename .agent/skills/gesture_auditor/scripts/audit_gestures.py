# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import os
import sys
import ast

# Whitelist of internal scripts that are exempt from docstring and direct hotkey checks
HELPER_WHITELIST = {
    "script_handleCommandKey",
    "script_cancelCommandPrefix"
}

# The main prefix trigger script (allowed to have a default gesture, but must have description/category)
PREFIX_TRIGGERS = {
    "script_triggerCommandPrefix"
}

def get_call_name(node):
    """Safely extracts function or method name from an AST Call node func."""
    if isinstance(node, ast.Name):
        return node.id
    elif isinstance(node, ast.Attribute):
        return node.attr
    return None

def find_app_module_managers(root_dir):
    """
    Dynamically scans all appModule files to discover manager associations from imports.
    E.g. if excel.py imports from appModules.boa_enhancements.excel_enhancements import manager,
    it automatically associates excel.py with excel_enhancements/manager.py.
    """
    app_module_dir = os.path.join(root_dir, "appModules")
    mappings = {}
    
    if not os.path.exists(app_module_dir):
        return mappings

    for file in os.listdir(app_module_dir):
        if file.endswith(".py"):
            filepath = os.path.join(app_module_dir, file)
            with open(filepath, "r", encoding="utf-8") as f:
                try:
                    tree = ast.parse(f.read())
                    for node in ast.walk(tree):
                        if isinstance(node, ast.ImportFrom) and node.module:
                            if "boa_enhancements" in node.module:
                                for name_alias in node.names:
                                    if name_alias.name == "manager":
                                        parts = node.module.split(".")
                                        try:
                                            idx = parts.index("boa_enhancements")
                                            if idx + 1 < len(parts):
                                                folder = parts[idx + 1]
                                                mappings[file] = f"{folder}/manager.py"
                                        except ValueError:
                                            pass
                except Exception as e:
                    print(f"Error reading imports in {file}: {e}")
    return mappings

class GestureVisitor(ast.NodeVisitor):
    def __init__(self, filepath):
        self.filepath = filepath
        self.gestures = {}  # script_name -> list of keys
        self.scripts = {}   # script_name -> {docstring, category, decorator_gestures, class, calls}
        self.class_category = None
        self.current_class = None

    def visit_ClassDef(self, node):
        old_class = self.current_class
        old_category = self.class_category
        self.current_class = node.name

        # Search for scriptCategory class attribute
        for stmt in node.body:
            if isinstance(stmt, ast.Assign):
                for target in stmt.targets:
                    if isinstance(target, ast.Name) and target.id == "scriptCategory":
                        val = stmt.value
                        if isinstance(val, ast.Call) and isinstance(val.func, ast.Name) and val.func.id == "_":
                            if val.args and isinstance(val.args[0], ast.Constant):
                                self.class_category = val.args[0].value
                            elif val.args and isinstance(val.args[0], ast.Str):
                                self.class_category = val.args[0].s
                        elif isinstance(val, ast.Constant):
                            self.class_category = val.value
                        elif isinstance(val, ast.Str):
                            self.class_category = val.s

        self.generic_visit(node)
        self.current_class = old_class
        self.class_category = old_category

    def visit_Assign(self, node):
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == "__gestures":
                if isinstance(node.value, ast.Dict):
                    for k, v in zip(node.value.keys, node.value.values):
                        key_str = None
                        val_str = None
                        if isinstance(k, ast.Constant):
                            key_str = k.value
                        elif isinstance(k, ast.Str):
                            key_str = k.s
                        if isinstance(v, ast.Constant):
                            val_str = v.value
                        elif isinstance(v, ast.Str):
                            val_str = v.s
                        if key_str and val_str:
                            if val_str not in self.gestures:
                                self.gestures[val_str] = []
                            self.gestures[val_str].append(key_str)
        self.generic_visit(node)

    def visit_FunctionDef(self, node):
        is_script = node.name.startswith("script_")
        if not is_script:
            for dec in node.decorator_list:
                if isinstance(dec, ast.Name) and dec.id == "script":
                    is_script = True
                elif isinstance(dec, ast.Call) and isinstance(dec.func, ast.Name) and dec.func.id == "script":
                    is_script = True

        if is_script:
            script_name = node.name
            docstring = ast.get_docstring(node)
            decorator_category = None
            decorator_gestures = []

            for dec in node.decorator_list:
                if isinstance(dec, ast.Call) and isinstance(dec.func, ast.Name) and dec.func.id == "script":
                    for kw in dec.keywords:
                        if kw.arg == "category":
                            val = kw.value
                            if isinstance(val, ast.Call) and isinstance(val.func, ast.Name) and val.func.id == "_":
                                if val.args and isinstance(val.args[0], ast.Constant):
                                    decorator_category = val.args[0].value
                                elif val.args and isinstance(val.args[0], ast.Str):
                                    decorator_category = val.args[0].s
                            elif isinstance(val, ast.Constant):
                                decorator_category = val.value
                            elif isinstance(val, ast.Str):
                                decorator_category = val.s
                        elif kw.arg == "gesture":
                            if isinstance(kw.value, ast.Constant):
                                decorator_gestures.append(kw.value.value)
                            elif isinstance(kw.value, ast.Str):
                                decorator_gestures.append(kw.value.s)
                        elif kw.arg == "gestures":
                            if isinstance(kw.value, ast.List):
                                for elt in kw.value.elts:
                                    if isinstance(elt, ast.Constant):
                                        decorator_gestures.append(elt.value)
                                    elif isinstance(elt, ast.Str):
                                        decorator_gestures.append(elt.s)

            # Find all function calls made inside this script
            calls = []
            for body_node in node.body:
                for sub_node in ast.walk(body_node):
                    if isinstance(sub_node, ast.Call):
                        name = get_call_name(sub_node.func)
                        if name:
                            calls.append(name)

            category = decorator_category or self.class_category
            self.scripts[script_name] = {
                "docstring": docstring,
                "category": category,
                "decorator_gestures": decorator_gestures,
                "class": self.current_class,
                "file": self.filepath,
                "calls": calls
            }
        self.generic_visit(node)

class ManagerVisitor(ast.NodeVisitor):
    def __init__(self):
        self.prefix_mappings = {} # key -> list of called function names

    def visit_FunctionDef(self, node):
        if node.name == "handle_prefix_command":
            for stmt in node.body:
                for child in ast.walk(stmt):
                    if isinstance(child, ast.If):
                        if isinstance(child.test, ast.Compare):
                            if isinstance(child.test.left, ast.Name) and child.test.left.id == "command_key":
                                for op, comparator in zip(child.test.ops, child.test.comparators):
                                    if isinstance(op, ast.Eq):
                                        key_str = None
                                        if isinstance(comparator, ast.Constant):
                                            key_str = comparator.value
                                        elif isinstance(comparator, ast.Str):
                                            key_str = comparator.s
                                        if key_str:
                                            # Gather calls in this condition's body (excluding orelse/elif branches)
                                            calls = []
                                            for body_node in child.body:
                                                for sub_node in ast.walk(body_node):
                                                    if isinstance(sub_node, ast.Call):
                                                        name = get_call_name(sub_node.func)
                                                        if name:
                                                            calls.append(name)
                                            self.prefix_mappings[key_str] = calls
        self.generic_visit(node)

def scan_repository(root_dir, app_module_pairs):
    all_scripts = {}
    all_gestures = {}  # gesture -> list of (file, script)
    app_prefix_mappings = {} # app_file_name -> {key -> list of calls}

    for root, dirs, files in os.walk(root_dir):
        if "__pycache__" in root or ".sconsign" in root:
            continue
        for file in files:
            if file.endswith(".py"):
                filepath = os.path.join(root, file)
                
                # Check if this file is a manager.py matched with an appModule
                matched_app = None
                for app_name, mgr_subpath in app_module_pairs.items():
                    normalized_subpath = mgr_subpath.replace("/", os.sep)
                    if filepath.endswith(normalized_subpath):
                        matched_app = app_name
                        break

                with open(filepath, "r", encoding="utf-8") as f:
                    try:
                        content = f.read()
                        tree = ast.parse(content)
                        
                        # Scan scripts and gestures
                        visitor = GestureVisitor(filepath)
                        visitor.visit(tree)
                        
                        # Extract dynamic docstrings (e.g. script_name.__doc__ = _("..."))
                        for child in ast.walk(tree):
                            if isinstance(child, ast.Assign):
                                for target in child.targets:
                                    if isinstance(target, ast.Attribute) and target.attr == "__doc__":
                                        if isinstance(target.value, ast.Name):
                                            s_name = target.value.id
                                            doc_val = None
                                            val = child.value
                                            if isinstance(val, ast.Call) and isinstance(val.func, ast.Name) and val.func.id == "_":
                                                if val.args and isinstance(val.args[0], ast.Constant):
                                                    doc_val = val.args[0].value
                                                elif val.args and isinstance(val.args[0], ast.Str):
                                                    doc_val = val.args[0].s
                                            elif isinstance(val, ast.Constant):
                                                doc_val = val.value
                                            elif isinstance(val, ast.Str):
                                                doc_val = val.s
                                            
                                            if doc_val and s_name in visitor.scripts:
                                                visitor.scripts[s_name]["docstring"] = doc_val
                        
                        for script_name, info in visitor.scripts.items():
                            all_scripts[script_name] = info
                            keys = visitor.gestures.get(script_name, []) + info["decorator_gestures"]
                            for key in keys:
                                if key not in all_gestures:
                                    all_gestures[key] = []
                                all_gestures[key].append((filepath, script_name))

                        # If manager, extract prefix key mappings
                        if matched_app:
                            mgr_visitor = ManagerVisitor()
                            mgr_visitor.visit(tree)
                            app_prefix_mappings[matched_app] = mgr_visitor.prefix_mappings

                    except Exception as e:
                        print(f"Error parsing {filepath}: {e}")

    return all_scripts, all_gestures, app_prefix_mappings

def check_gestures(all_scripts, all_gestures, app_prefix_mappings):
    errors = []
    warnings = []

    # 1. Check for duplicate gestures
    for gesture, bindings in all_gestures.items():
        if len(bindings) > 1:
            files = [b[0] for b in bindings]
            scripts = [b[1] for b in bindings]
            errors.append(f"Conflict: Gesture '{gesture}' is mapped to multiple scripts: {scripts} in files {files}")

    # 2. Validate script exposure and categorization
    for script_name, info in all_scripts.items():
        if script_name in HELPER_WHITELIST:
            continue

        expected_cat = "BOA (Better Office Accessibility)"
        if info["category"] != expected_cat:
            errors.append(f"Category Mismatch: Script '{script_name}' in {info['file']} has category '{info['category']}', expected '{expected_cat}'")

        if not info["docstring"] and script_name not in PREFIX_TRIGGERS:
            errors.append(f"Missing Docstring: Script '{script_name}' in {info['file']} has no description/docstring.")

        if script_name not in PREFIX_TRIGGERS:
            keys_defined = [g for g, binds in all_gestures.items() if any(b[1] == script_name for b in binds)]
            if keys_defined:
                errors.append(f"Invalid Default Key: Script '{script_name}' in {info['file']} defines default key bindings: {keys_defined}. Custom scripts must be empty by default so users choose their own hotkeys.")

    # 3. Dynamic Prefix Mode Exposure Audit (Zero Hardcoding)
    for app_file, key_mappings in app_prefix_mappings.items():
        # Filter scripts belonging to this specific appModule file
        app_scripts = {s: info for s, info in all_scripts.items() if os.path.basename(info["file"]) == app_file}
        
        for key, called_funcs in key_mappings.items():
            # Skip digital slot assignments and shift+digit assignments
            if len(key) == 1 and key.isdigit():
                continue
            if key.startswith("shift+"):
                continue

            # Verify that at least one script in the appModule calls one of the functions invoked by this key
            is_exposed = False
            for func_name in called_funcs:
                # Find if any script in this appModule calls this function
                for script_name, script_info in app_scripts.items():
                    if func_name in script_info["calls"]:
                        is_exposed = True
                        break
                if is_exposed:
                    break

            if not is_exposed:
                errors.append(f"Missing Exposure: {app_file} prefix command '{key}' (which executes '{called_funcs}') has no corresponding standalone script calling these functions.")

    return errors, warnings

def main():
    addon_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "addon"))
    print(f"Scanning BOA addon directory: {addon_dir}")
    
    # Dynamically discover AppModule -> Manager associations via imports
    app_module_pairs = find_app_module_managers(addon_dir)
    print(f"Discovered appModule-manager pairs: {app_module_pairs}")
    
    all_scripts, all_gestures, app_prefix_mappings = scan_repository(addon_dir, app_module_pairs)
    
    errors, warnings = check_gestures(all_scripts, all_gestures, app_prefix_mappings)

    print("\n" + "="*50)
    print("BOA KEYSTROKE & GESTURE AUDIT REPORT (FULLY AUTOMATED)")
    print("="*50)

    print(f"\nTotal Discovered Scripts: {len(all_scripts)}")
    for name, info in sorted(all_scripts.items()):
        status = "Internal (Hidden)" if name in HELPER_WHITELIST else f"Category: {info['category']}"
        print(f" - {name} ({status})")

    for app, key_mappings in app_prefix_mappings.items():
        if key_mappings:
            print(f"\nPrefix Commands Monitored in {app}: {list(key_mappings.keys())}")

    if warnings:
        print("\n" + "-"*50)
        print(f"WARNINGS ({len(warnings)})")
        print("-"*50)
        for w in warnings:
            print(f" [WARNING] {w}")

    if errors:
        print("\n" + "-"*50)
        print(f"ERRORS / CONFLICTS ({len(errors)})")
        print("-"*50)
        for e in errors:
            print(f" [ERROR] {e}")
        print("\nAudit Failed! Please resolve the errors above.")
        sys.exit(1)
    else:
        print("\nAudit Passed! All gestures are conflict-free, correctly grouped, and properly exposed dynamically.")
        sys.exit(0)

if __name__ == "__main__":
    main()
