import matplotlib.pyplot as plt
import numpy as np

# Data for global diabetes prevalence (2021 projection)
years = ['2011', '2013', '2015', '2017', '2019', '2021', '2030', '2045']
global_cases = [366, 382, 415, 425, 463, 537, 643, 783]  # in millions

# Indian diabetes prevalence trends
india_years = ['2000', '2005', '2010', '2015', '2020', '2021', '2030', '2045']
india_prevalence = [5.5, 6.5, 8.3, 9.8, 10.8, 11.4, 14.5, 16.0]
india_cases = [31.7, 40.9, 61.3, 69.2, 88.9, 101.2, 140.2, 160.0]  # in millions

# Create subplots
fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
fig.suptitle('DIABETES EPIDEMIOLOGY: GLOBAL AND INDIA', fontsize=16, fontweight='bold')

# Global diabetes cases over time
ax1.plot(years, global_cases, marker='o', linewidth=3, markersize=8, color='#1976D2')
ax1.fill_between(years, global_cases, alpha=0.3, color='#1976D2')
ax1.set_title('Global Diabetes Cases (Millions)', fontsize=14, fontweight='bold')
ax1.set_ylabel('Cases (Millions)')
ax1.grid(True, alpha=0.3)
for i, v in enumerate(global_cases):
    ax1.text(i, v + 10, f'{v}', ha='center', va='bottom', fontweight='bold')

# India prevalence over time
ax2.plot(india_years, india_prevalence, marker='s', linewidth=3, markersize=8, color='#388E3C')
ax2.fill_between(india_years, india_prevalence, alpha=0.3, color='#388E3C')
ax2.set_title('India: Diabetes Prevalence (%)', fontsize=14, fontweight='bold')
ax2.set_ylabel('Prevalence (%)')
ax2.grid(True, alpha=0.3)
for i, v in enumerate(india_prevalence):
    ax2.text(i, v + 0.3, f'{v}%', ha='center', va='bottom', fontweight='bold')

# India cases over time
ax3.bar(india_years, india_cases, color='#F57C00', alpha=0.8)
ax3.set_title('India: Diabetes Cases (Millions)', fontsize=14, fontweight='bold')
ax3.set_ylabel('Cases (Millions)')
ax3.grid(True, alpha=0.3, axis='y')
for i, v in enumerate(india_cases):
    ax3.text(i, v + 2, f'{v}', ha='center', va='bottom', fontweight='bold')

# Regional comparison
regions = ['South Asia', 'East Asia &\nPacific', 'North\nAmerica', 'Western\nEurope', 'Middle East &\nN Africa', 'South &\nCentral America', 'Sub-Saharan\nAfrica']
prevalence = [8.8, 8.4, 11.7, 5.4, 14.0, 9.9, 3.3]
colors = ['#4CAF50', '#2196F3', '#FF9800', '#9C27B0', '#F44336', '#795548', '#607D8B']

bars = ax4.bar(regions, prevalence, color=colors, alpha=0.8)
ax4.set_title('Regional Diabetes Prevalence (2021)', fontsize=14, fontweight='bold')
ax4.set_ylabel('Prevalence (%)')
ax4.tick_params(axis='x', rotation=45)
ax4.grid(True, alpha=0.3, axis='y')

for bar, value in zip(bars, prevalence):
    ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
             f'{value}%', ha='center', va='bottom', fontweight='bold', fontsize=10)

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/epidemiology_chart.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
plt.close()

print("Epidemiology chart saved as epidemiology_chart.png")
