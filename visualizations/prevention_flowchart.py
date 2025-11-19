import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Create figure and axis
fig, ax = plt.subplots(figsize=(14, 16))
ax.set_xlim(0, 14)
ax.set_ylim(0, 16)
ax.axis('off')

# Define colors and styles
colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#e67e22']
decision_color = '#3498db'
action_color = '#ecf0f1'
primary_color = '#27ae60'
secondary_color = '#e67e22'
text_color = '#2c3e50'

# Title
ax.text(7, 15.5, 'DIABETES PREVENTION STRATEGY', ha='center', va='center',
        fontsize=18, fontweight='bold', color=text_color)

# Primary Prevention (High Risk)
ax.add_patch(patches.FancyBboxPatch((4, 13.5), 6, 1, boxstyle="round,pad=0.2",
                                   facecolor=primary_color, edgecolor=primary_color))
ax.text(7, 14, 'PRIMARY PREVENTION', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(7, 13.8, 'Individuals at High Risk of Developing Diabetes', ha='center', va='center', fontsize=9, color='white')

# Decision point
ax.add_patch(patches.FancyBboxPatch((5.5, 12), 5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=decision_color, edgecolor=decision_color))
ax.text(8, 12.4, 'High Risk Identified?', ha='center', va='center', fontsize=11, fontweight='bold', color='white')

# Arrows
ax.arrow(8, 12, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# High risk identified
ax.add_patch(patches.FancyBboxPatch((8.5, 10.5), 4, 1, boxstyle="round,pad=0.1",
                                   facecolor=action_color, edgecolor=colors[2]))
ax.text(10.5, 11, 'Lifestyle Intervention\n(DPP Model)', ha='center', va='center', fontsize=10, fontweight='bold')

# Lifestyle components
y_pos = 9.5
lifestyle_items = [
    '• Weight loss 5-7% of body weight',
    '• Physical activity 150 min/week',
    '• Dietary modification',
    '• 3-year maintenance program'
]
for item in lifestyle_items:
    ax.text(10.5, y_pos, item, ha='center', va='center', fontsize=8)
    y_pos -= 0.5

ax.arrow(10.5, 9, 0, -0.7, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Pharmacological option
ax.add_patch(patches.FancyBboxPatch((8.5, 7), 4, 1, boxstyle="round,pad=0.1",
                                   facecolor=action_color, edgecolor=colors[3]))
ax.text(10.5, 7.5, 'Pharmacologic Prevention', ha='center', va='center', fontsize=10, fontweight='bold')

ax.text(10.5, 7, 'Metformin 850mg BID', ha='center', va='center', fontsize=8)
ax.text(10.5, 6.5, '(if BMI ≥35 and age <60)', ha='center', va='center', fontsize=7)

ax.arrow(10.5, 6, 0, -0.7, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Secondary Prevention (Prediabetes)
ax.add_patch(patches.FancyBboxPatch((0.5, 4.5), 6, 1, boxstyle="round,pad=0.2",
                                   facecolor=secondary_color, edgecolor=secondary_color))
ax.text(3.5, 5, 'SECONDARY PREVENTION', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(3.5, 4.8, 'Prediabetes Management', ha='center', va='center', fontsize=9, color='white')

# Prediabetes screening
ax.add_patch(patches.FancyBboxPatch((1.5, 3), 4, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_color, edgecolor=colors[0]))
ax.text(3.5, 3.4, 'Screening & Diagnosis', ha='center', va='center', fontsize=10, fontweight='bold')

ax.text(3.5, 2.8, 'IFG: FPG 100-125 mg/dL', ha='center', va='center', fontsize=8)
ax.text(3.5, 2.4, 'IGT: 2hPG 140-199 mg/dL', ha='center', va='center', fontsize=8)
ax.text(3.5, 2, 'Conversion rate: 5-10%/year', ha='center', va='center', fontsize=7)

ax.arrow(3.5, 2.8, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Prediabetes management
ax.add_patch(patches.FancyBboxPatch((1.5, 1), 4, 1, boxstyle="round,pad=0.1",
                                   facecolor=action_color, edgecolor=colors[4]))
ax.text(3.5, 1.5, 'Prediabetes Management', ha='center', va='center', fontsize=10, fontweight='bold')

ax.text(3.5, 1, 'Weight loss 5-10%', ha='center', va='center', fontsize=8)
ax.text(3.5, 0.6, 'Exercise + Diet', ha='center', va='center', fontsize=8)
ax.text(3.5, 0.2, 'Metformin optional', ha='center', va='center', fontsize=7)

# Tertiary Prevention (Complications)
ax.add_patch(patches.FancyBboxPatch((8, 4.5), 5.5, 1, boxstyle="round,pad=0.2",
                                   facecolor=colors[6], edgecolor=colors[6]))
ax.text(10.75, 5, 'TERTIARY PREVENTION', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(10.75, 4.8, 'Preventing Complications & Disability', ha='center', va='center', fontsize=9, color='white')

# Complications management
ax.add_patch(patches.FancyBboxPatch((9, 3), 4, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_color, edgecolor=colors[1]))
ax.text(11, 3.4, 'Microvascular Prevention', ha='center', va='center', fontsize=10, fontweight='bold')

ax.text(11, 2.8, 'Retinopathy screening', ha='center', va='center', fontsize=8)
ax.text(11, 2.4, 'Nephropathy monitoring', ha='center', va='center', fontsize=8)
ax.text(11, 2, 'Neuropathy assessment', ha='center', va='center', fontsize=7)

ax.arrow(11, 2.8, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

ax.add_patch(patches.FancyBboxPatch((9, 1), 4, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_color, edgecolor=colors[5]))
ax.text(11, 1.5, 'Macrovascular Prevention', ha='center', va='center', fontsize=10, fontweight='bold')

ax.text(11, 1, 'Blood pressure <130/80', ha='center', va='center', fontsize=8)
ax.text(11, 0.6, 'Statin therapy', ha='center', va='center', fontsize=8)
ax.text(11, 0.2, 'Antiplatelet therapy', ha='center', va='center', fontsize=7)

# Risk factor assessment
ax.add_patch(patches.FancyBboxPatch((4, 7.5), 4, 1.2, boxstyle="round,pad=0.1",
                                   facecolor=colors[0], edgecolor=colors[0]))
ax.text(6, 8.1, 'RISK FACTOR ASSESSMENT', ha='center', va='center', fontsize=12, fontweight='bold', color='white')

risk_factors = [
    '• Family history',
    '• BMI >25 kg/m²'
    '• Age >45 years',
    '• Sedentary lifestyle',
    '• Hypertension',
    '• Dyslipidemia'
]

y_pos = 7.5
for factor in risk_factors:
    ax.text(6, y_pos, factor, ha='center', va='center', fontsize=7, color='white')
    y_pos -= 0.2

ax.arrow(6, 7.5, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Results/Outcomes box
ax.add_patch(patches.FancyBboxPatch((5, 9.5), 4, 1.5, boxstyle="round,pad=0.1",
                                   facecolor=colors[3], edgecolor=colors[3]))
ax.text(7, 10.3, 'OUTCOMES', ha='center', va='center', fontsize=13, fontweight='bold', color='white')

outcomes = [
    '• 58% reduction in incidence',
    '• 31% reduction with metformin',
    '• Cost-effective prevention',
    '• Long-term benefits'
]

y_pos = 10
for outcome in outcomes:
    ax.text(7, y_pos, outcome, ha='center', va='center', fontsize=8, color='white')
    y_pos -= 0.2

ax.arrow(7, 9.5, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Connecting lines and arrows
# From risk assessment to primary prevention
ax.arrow(6, 8.5, -1.5, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')

# From secondary to tertiary
ax.arrow(3.5, 4.5, 5, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(8, 4.7, 'Progression to diabetes', ha='center', va='center', fontsize=8)

# Legend
ax.add_patch(patches.FancyBboxPatch((1, 14), 3, 1.2, boxstyle="round,pad=0.1",
                                   facecolor='white', edgecolor='black', alpha=0.9))
ax.text(2.5, 14.5, 'LEGEND', ha='center', va='center', fontsize=11, fontweight='bold')
ax.text(2.5, 14, '🟢 Primary Prevention', ha='center', va='center', fontsize=8)
ax.text(2.5, 13.6, '🟠 Secondary Prevention', ha='center', va='center', fontsize=8)
ax.text(2.5, 13.2, '🟤 Tertiary Prevention', ha='center', va='center', fontsize=8)

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/prevention_flowchart.png', dpi=300, bbox_inches='tight', pad_inches=1)
plt.close()

print("Prevention flowchart saved as prevention_flowchart.png")
