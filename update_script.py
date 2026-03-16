import traceback

def main():
    try:
        with open('payroll_automation.py', 'r', encoding='utf-8') as f:
            code = f.read()

        # Remove red borders and replace with regular thin borders
        code = code.replace(", color=\"FF0000\"", "")
        code = code.replace(", color='FF0000'", "")
        code = code.replace("red_thin_border", "thin_border")
        code = code.replace("red_side", "Side(style='thin')")

        with open('payroll_automation.py', 'w', encoding='utf-8') as f:
            f.write(code)
        print("Successfully updated colors in payroll_automation.py")
    except Exception as e:
        print("Error:", e)
        traceback.print_exc()

if __name__ == '__main__':
    main()
