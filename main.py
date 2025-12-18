import keyboard
from speech_to_text import listen_once
from gemini_ai import interpret_command
import excel_actions as excel
import time


def main():
    print("Hold SPACE to speak your Excel command...")

    while True:

        text = listen_once()

        if not text:
            print("⚠ No speech detected")
            continue

        # Convert speech → JSON command
        cmd = interpret_command(text)   
        action = cmd.get("action", "unknown")

        # ------------- BASIC ACTIONS ------------- #

        if action == "write":
            if "cell" in cmd and "value" in cmd:
                excel.write_cell(cmd["cell"], cmd["value"])
                print(f"✔ Wrote '{cmd['value']}' in {cmd['cell']}")
            else:
                print("⚠ Missing 'cell' or 'value' field")

        elif action == "delete_cell":
            if "cell" in cmd:
                excel.delete_cell(cmd["cell"])
                print(f"✔ Cleared {cmd['cell']}")
            else:
                print("⚠ Missing 'cell' field")

        elif action == "insert_row":
            if "row" in cmd:
                excel.insert_row(cmd["row"])
                print(f"✔ Inserted row {cmd['row']}")
            else:
                print("⚠ Missing 'row' field")

        elif action == "insert_column":
            if "column" in cmd:
                excel.insert_column(cmd["column"])
                print(f"✔ Inserted column {cmd['column']}")
            else:
                print("⚠ Missing 'column' field")


        # ------------- ADVANCED ACTIONS ------------- #

        elif action == "sum_column":
            if "column" in cmd:
                result_cell = excel.sum_column(cmd["column"])
                print(f"✔ SUM result placed at {result_cell}")
            else:
                print("⚠ Missing 'column' field")

        elif action == "sort_column":
            col = cmd.get("column")
            order = cmd.get("order", "asc")
            if col:
                excel.sort_column(col, order)
                print(f"✔ Sorted column {col} ({order})")
            else:
                print("⚠ Missing 'column' field")

        elif action == "format_bold":
            if "column" in cmd:
                excel.format_bold(cmd["column"])
                print(f"✔ Bold formatting applied to {cmd['column']}")
            else:
                print("⚠ Missing 'column' field")

        elif action == "filter_values":
            col = cmd.get("column")
            cond = cmd.get("condition")
            if col and cond:
                excel.filter_values(col, cond)
                print(f"✔ Filter applied on {col} with condition '{cond}'")
            else:
                print("⚠ Missing 'column' or 'condition'")

        elif action == "create_chart":
            x = cmd.get("x_column")
            y = cmd.get("y_column")
            if x and y:
                excel.create_chart(x, y)
                print(f"✔ Chart created using {x} vs {y}")
            else:
                print("⚠ Missing 'x_column' or 'y_column'")

        elif action == "run_regression":
            x = cmd.get("x_column")
            y = cmd.get("y_column")
            if x and y:
                excel.run_regression(x, y)
                print(f"✔ Regression executed with {x} → {y}")
            else:
                print("⚠ Missing 'x_column' or 'y_column'")


        # ------------- UNKNOWN ------------- #
        else:
            print("❌ Unknown or unsupported command")

        print("\nReady for next command...\n")


if __name__ == "__main__":
    main()
