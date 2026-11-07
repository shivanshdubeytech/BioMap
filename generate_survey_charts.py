import matplotlib.pyplot as plt
import numpy as np

# Global typography and aesthetic styling
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Helvetica']

# Color tokens aligned with report palette
DARK_EMERALD = '#1B4332'
FOREST_GREEN = '#2D6A4F'
TERRACOTTA = '#C05621'
SLATE_BLUE = '#1E40AF'
ACCENT_RED = '#DC2626'
TEXT_DARK = '#1A202C'
TEXT_MUTED = '#4A5568'
GRID_COLOR = '#94A3B8'

print("Generating high-resolution survey charts with enhanced typography...")

# ----------------------------------------------------
# 1. Figure 4.5: Where is Free-Living Wildlife Seen?
# ----------------------------------------------------
fig, ax = plt.subplots(figsize=(5.0, 3.4), dpi=300)
categories_4_5 = [
    'Botanical Grove (Z1)',
    'Freshwater Wetland (Z2)',
    'Promenade Shrublands (Z3)',
    'Perimeter Trees (Z4)',
    'Visitor Lawns',
    'Other Areas'
]
counts_4_5 = [26, 21, 17, 8, 7, 2]
y_pos_4_5 = np.arange(len(categories_4_5))

bars_4_5 = ax.barh(y_pos_4_5, counts_4_5, color=DARK_EMERALD, height=0.65, edgecolor='#081C15', linewidth=0.8)
ax.set_yticks(y_pos_4_5)
ax.set_yticklabels(categories_4_5, fontsize=10.5, fontweight='bold', color=TEXT_DARK)
ax.invert_yaxis()
ax.set_xlabel('Survey Count (N = 84)', fontsize=10, fontweight='bold', color=TEXT_MUTED)
ax.set_title('Where Is Free-Living Wildlife Seen?', fontsize=12, fontweight='bold', loc='left', color=DARK_EMERALD, pad=10)

for bar in bars_4_5:
    w = bar.get_width()
    ax.text(w + 0.6, bar.get_y() + bar.get_height()/2, f'{int(w)}', va='center', fontsize=10.5, fontweight='bold', color=DARK_EMERALD)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#718096')
ax.spines['bottom'].set_color('#718096')
ax.set_xlim(0, 31)
ax.xaxis.grid(True, linestyle='--', alpha=0.35, color=GRID_COLOR)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig('fig_4_5.png', bbox_inches='tight')
plt.close()
print("[OK] fig_4_5.png generated")

# ----------------------------------------------------
# 2. Figure 4.6: What Threatens Free-Living Flora & Avifauna?
# ----------------------------------------------------
fig, ax = plt.subplots(figsize=(5.0, 3.4), dpi=300)
categories_4_6 = [
    'Arbitrary branch pruning',
    'Preserved in inner grove',
    'Disturbed by traffic noise',
    'Nesting holes lost',
    'Not sure / Indifferent'
]
counts_4_6 = [23, 16, 12, 9, 3]
y_pos_4_6 = np.arange(len(categories_4_6))

bars_4_6 = ax.barh(y_pos_4_6, counts_4_6, color=TERRACOTTA, height=0.65, edgecolor='#7B341E', linewidth=0.8)
ax.set_yticks(y_pos_4_6)
ax.set_yticklabels(categories_4_6, fontsize=10.5, fontweight='bold', color=TEXT_DARK)
ax.invert_yaxis()
ax.set_xlabel('Survey Count (N = 84)', fontsize=10, fontweight='bold', color=TEXT_MUTED)
ax.set_title('What Threatens Free-Living Flora & Fauna?', fontsize=12, fontweight='bold', loc='left', color=TERRACOTTA, pad=10)

for bar in bars_4_6:
    w = bar.get_width()
    ax.text(w + 0.6, bar.get_y() + bar.get_height()/2, f'{int(w)}', va='center', fontsize=10.5, fontweight='bold', color=TERRACOTTA)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#718096')
ax.spines['bottom'].set_color('#718096')
ax.set_xlim(0, 27)
ax.xaxis.grid(True, linestyle='--', alpha=0.35, color=GRID_COLOR)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig('fig_4_6.png', bbox_inches='tight')
plt.close()
print("[OK] fig_4_6.png generated")

# ----------------------------------------------------
# 3. Figure 4.7: Main Barriers to Biodiversity Literacy
# ----------------------------------------------------
fig, ax = plt.subplots(figsize=(5.0, 3.4), dpi=300)
categories_4_7 = [
    'No digital map of trees & fauna',
    'Signboards only on caged animals',
    'Lack of time during visit',
    'Do not know species names',
    'Unaware of heritage value',
    'No barrier observed'
]
counts_4_7 = [18, 15, 12, 11, 8, 3]
y_pos_4_7 = np.arange(len(categories_4_7))

bars_4_7 = ax.barh(y_pos_4_7, counts_4_7, color=SLATE_BLUE, height=0.65, edgecolor='#1E3A8A', linewidth=0.8)
ax.set_yticks(y_pos_4_7)
ax.set_yticklabels(categories_4_7, fontsize=10.5, fontweight='bold', color=TEXT_DARK)
ax.invert_yaxis()
ax.set_xlabel('Survey Count (N = 84)', fontsize=10, fontweight='bold', color=TEXT_MUTED)
ax.set_title('Main Barriers to Biodiversity Literacy', fontsize=12, fontweight='bold', loc='left', color=SLATE_BLUE, pad=10)

for bar in bars_4_7:
    w = bar.get_width()
    ax.text(w + 0.5, bar.get_y() + bar.get_height()/2, f'{int(w)}', va='center', fontsize=10.5, fontweight='bold', color=SLATE_BLUE)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#718096')
ax.spines['bottom'].set_color('#718096')
ax.set_xlim(0, 22)
ax.xaxis.grid(True, linestyle='--', alpha=0.35, color=GRID_COLOR)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig('fig_4_7.png', bbox_inches='tight')
plt.close()
print("[OK] fig_4_7.png generated")

# ----------------------------------------------------
# 4. Figure 4.8 & 4.9: Side-by-Side Pie Charts
# ----------------------------------------------------
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(5.6, 3.4), dpi=300)

labels1 = ['Yes\n(60%)', 'Maybe\n(30%)', 'No\n(10%)']
sizes1 = [18, 9, 3]
colors1 = ['#1B4332', '#2563EB', '#DC2626']
wedges1, texts1, autotexts1 = ax1.pie(
    sizes1, labels=labels1, colors=colors1, autopct='%1.0f%%', startangle=140,
    textprops={'fontsize': 10, 'fontweight': 'bold', 'color': TEXT_DARK},
    pctdistance=0.65
)
for at in autotexts1:
    at.set_color('white')
    at.set_fontweight('bold')
    at.set_fontsize(10.5)
ax1.set_title('Would an App / Digital\nMap Help?', fontsize=11.5, fontweight='bold', color='#1B4332', pad=8)

labels2 = ['Definitely\n(53%)', 'Likely\n(33%)', 'Unsure\n(14%)']
sizes2 = [16, 10, 4]
colors2 = ['#2D6A4F', '#D97706', '#94A3B8']
wedges2, texts2, autotexts2 = ax2.pie(
    sizes2, labels=labels2, colors=colors2, autopct='%1.0f%%', startangle=140,
    textprops={'fontsize': 10, 'fontweight': 'bold', 'color': TEXT_DARK},
    pctdistance=0.65
)
for at in autotexts2:
    at.set_color('white')
    at.set_fontweight('bold')
    at.set_fontsize(10.5)
ax2.set_title('Would Citizens Actively\nUse BioMap?', fontsize=11.5, fontweight='bold', color='#2D6A4F', pad=8)

plt.tight_layout()
plt.savefig('fig_4_8_4_9.png', bbox_inches='tight')
plt.close()
print("[OK] fig_4_8_4_9.png generated")

# ----------------------------------------------------
# 5. Figure 4.10: What Info Should Species Listing Include?
# ----------------------------------------------------
fig, ax = plt.subplots(figsize=(5.0, 3.4), dpi=300)
categories_4_10 = [
    'GPS on Map',
    'High-Res Photo',
    'Common & Sci Name',
    'Flowering Season',
    'Ecological Role',
    'IUCN Status'
]
counts_4_10 = [21, 19, 16, 13, 12, 10]
y_pos_4_10 = np.arange(len(categories_4_10))

bars_4_10 = ax.barh(y_pos_4_10, counts_4_10, color=DARK_EMERALD, height=0.65, edgecolor='#081C15', linewidth=0.8)
ax.set_yticks(y_pos_4_10)
ax.set_yticklabels(categories_4_10, fontsize=10.5, fontweight='bold', color=TEXT_DARK)
ax.invert_yaxis()
ax.set_xlabel('Survey Count (N = 84)', fontsize=10, fontweight='bold', color=TEXT_MUTED)
ax.set_title('What Info Should Listing Include?', fontsize=11.5, fontweight='bold', loc='left', color=DARK_EMERALD, pad=10)

for bar in bars_4_10:
    w = bar.get_width()
    ax.text(w + 0.5, bar.get_y() + bar.get_height()/2, f'{int(w)}', va='center', fontsize=10.5, fontweight='bold', color=DARK_EMERALD)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#718096')
ax.spines['bottom'].set_color('#718096')
ax.set_xlim(0, 25)
ax.xaxis.grid(True, linestyle='--', alpha=0.35, color=GRID_COLOR)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig('fig_4_10.png', bbox_inches='tight')
plt.close()
print("[OK] fig_4_10.png generated")

# ----------------------------------------------------
# 6. Figure 4.11 & 4.12 Combined High-Res Panel
# ----------------------------------------------------
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(5.0, 3.4), dpi=300, gridspec_kw={'height_ratios': [1.2, 1.0]})

# Top: Fig 4.11
cats_11 = ['Combined (All)', 'By Zone (Z1-Z4)', 'By Taxon Group', 'By IUCN Status']
cts_11 = [17, 7, 4, 2]
y11 = np.arange(len(cats_11))
bars11 = ax1.barh(y11, cts_11, color=ACCENT_RED, height=0.6, edgecolor='#991B1B', linewidth=0.8)
ax1.set_yticks(y11)
ax1.set_yticklabels(cats_11, fontsize=9.5, fontweight='bold', color=TEXT_DARK)
ax1.invert_yaxis()
ax1.set_title('Fig. 4.11: How to Discover Species?', fontsize=10.5, fontweight='bold', loc='left', color=ACCENT_RED)
for bar in bars11:
    w = bar.get_width()
    ax1.text(w + 0.4, bar.get_y() + bar.get_height()/2, f'{int(w)}', va='center', fontsize=9.5, fontweight='bold', color=ACCENT_RED)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_xlim(0, 20)
ax1.xaxis.grid(True, linestyle='--', alpha=0.3, color=GRID_COLOR)

# Bottom: Fig 4.12
labels12 = ['Yes (57%)', 'Sometimes (30%)', 'No (13%)']
sizes12 = [17, 9, 4]
colors12 = [DARK_EMERALD, '#2563EB', '#CBD5E1']
wedges, texts, autotexts = ax2.pie(
    sizes12, labels=labels12, colors=colors12, autopct='%1.0f%%', startangle=140,
    textprops={'fontsize': 9, 'fontweight': 'bold', 'color': TEXT_DARK},
    pctdistance=0.65
)
for at in autotexts:
    at.set_color('white')
    at.set_fontweight('bold')
    at.set_fontsize(9.5)
ax2.set_title('Fig. 4.12: Is QR Placard on Trees Useful?', fontsize=10.5, fontweight='bold', color=DARK_EMERALD, pad=4)

plt.tight_layout()
plt.savefig('fig_4_11_4_12.png', bbox_inches='tight')
plt.close()
print("[OK] fig_4_11_4_12.png generated")

# Standalone 4.11 and 4.12 as fallbacks
fig, ax = plt.subplots(figsize=(5.0, 2.2), dpi=300)
bars11_s = ax.barh(y11, cts_11, color=ACCENT_RED, height=0.6, edgecolor='#991B1B', linewidth=0.8)
ax.set_yticks(y11)
ax.set_yticklabels(cats_11, fontsize=9.5, fontweight='bold', color=TEXT_DARK)
ax.invert_yaxis()
ax.set_title('How to Discover Species?', fontsize=10.5, fontweight='bold', loc='left', color=ACCENT_RED)
for bar in bars11_s:
    w = bar.get_width()
    ax.text(w + 0.4, bar.get_y() + bar.get_height()/2, f'{int(w)}', va='center', fontsize=9.5, fontweight='bold', color=ACCENT_RED)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.set_xlim(0, 20)
ax.xaxis.grid(True, linestyle='--', alpha=0.3, color=GRID_COLOR)
plt.tight_layout()
plt.savefig('fig_4_11.png', bbox_inches='tight')
plt.close()

fig, ax = plt.subplots(figsize=(3.5, 2.2), dpi=300)
wedges, texts, autotexts = ax.pie(
    sizes12, labels=labels12, colors=colors12, autopct='%1.0f%%', startangle=140,
    textprops={'fontsize': 9, 'fontweight': 'bold', 'color': TEXT_DARK},
    pctdistance=0.65
)
for at in autotexts:
    at.set_color('white')
    at.set_fontweight('bold')
    at.set_fontsize(9.5)
ax.set_title('Is QR Placard on Trees Useful?', fontsize=10.5, fontweight='bold', color=DARK_EMERALD, pad=4)
plt.tight_layout()
plt.savefig('fig_4_12.png', bbox_inches='tight')
plt.close()

print("All survey charts regenerated successfully with large bold fonts!")
