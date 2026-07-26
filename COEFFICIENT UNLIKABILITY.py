import ipywidgets as widgets
from IPython.display import display, clear_output

# Create input fields for a, b, c, d
style = {'description_width': 'initial'}
a_input = widgets.IntText(value=45, description='a (TP):', style=style)
b_input = widgets.IntText(value=5,  description='b (FP):', style=style)
c_input = widgets.IntText(value=5,  description='c (FN):', style=style)
d_input = widgets.IntText(value=45, description='d (TN):', style=style)

output = widgets.Output()

def calculate_cu(*args):
    a = a_input.value
    b = b_input.value
    c = c_input.value
    d = d_input.value
    
    with output:
        clear_output()
        
        # Validation for denominators
        if (a + c) == 0 or (b + d) == 0:
            print("⚠️ Error: (a + c) and (b + d) must be greater than zero.")
            return

        # --- Section 1: Positive Group CU% (CU1) ---
        p1_1 = a / (a + c)
        p2_1 = 1 - p1_1
        cu1 = 2 * p1_1 * p2_1 * 100

        # --- Section 2: Negative Group CU% (CU2) ---
        p1_2 = d / (b + d)
        p2_2 = 1 - p1_2
        cu2 = 2 * p1_2 * p2_2 * 100

        # --- Section 3: Average CU% ---
        avg_cu = (cu1 + cu2) / 2

        # Acceptance Criteria Check
        passed = avg_cu <= 10
        status_text = "ACCEPTED (<= 10%)" if passed else "REJECTED (> 10%)"
        status_symbol = "✅" if passed else "❌"

        # Display Formatted Results
        print("=" * 48)
        print("       CUMULATIVE UNCERTAINTY (CU%) ANALYSIS    ")
        print("=" * 48)
        print(f" Positive Group (CU1%):")
        print(f"   • p1 = {p1_1:.4f} | p2 = {p2_1:.4f}")
        print(f"   • CU1% = {cu1:.2f}%")
        print("-" * 48)
        print(f" Negative Group (CU2%):")
        print(f"   • p1 = {p1_2:.4f} | p2 = {p2_2:.4f}")
        print(f"   • CU2% = {cu2:.2f}%")
        print("-" * 48)
        print(f" Average CU%        : {avg_cu:.2f}%")
        print(f" Evaluation Status  : {status_symbol} {status_text}")
        print("=" * 48)

# Attach event listeners to auto-update on input change
a_input.observe(calculate_cu, names='value')
b_input.observe(calculate_cu, names='value')
c_input.observe(calculate_cu, names='value')
d_input.observe(calculate_cu, names='value')

# Layout display
ui = widgets.VBox([
    widgets.HTML("<h3>Enter Matrix Values (a, b, c, d)</h3>"),
    widgets.HBox([a_input, b_input]),
    widgets.HBox([c_input, d_input]),
])

display(ui, output)

# Run once initially with default values
calculate_cu()