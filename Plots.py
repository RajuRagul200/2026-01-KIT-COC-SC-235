import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import make_interp_spline

# Font settings
plt.rcParams["font.family"] = "Times New Roman"
plt.rcParams["font.size"] = 16
plt.rcParams["font.weight"] = "bold"

# Time points
x = np.array([1, 2, 3])
x_labels = ["T1", "T2", "T3"]

# Mean values
cultural_awareness = np.array([3.12, 4.28, 4.05])
cultural_attachment = np.array([3.18, 4.35, 4.20])
continuity_intention = np.array([3.22, 4.40, 4.18])

# Smooth x-axis
x_smooth = np.linspace(x.min(), x.max(), 300)

# Quadratic spline (k=2) → works for 3 points
ca_smooth = make_interp_spline(x, cultural_awareness, k=2)(x_smooth)
att_smooth = make_interp_spline(x, cultural_attachment, k=2)(x_smooth)
ci_smooth = make_interp_spline(x, continuity_intention, k=2)(x_smooth)

# Plot
plt.figure(figsize=(8, 6))

plt.plot(x_smooth, ca_smooth, linewidth=2, label="Cultural Awareness")
plt.plot(x_smooth, att_smooth, linewidth=2, label="Cultural Attachment")
plt.plot(x_smooth, ci_smooth, linewidth=2, label="Continuity Intention")

plt.xticks(x, x_labels, fontweight="bold")
plt.xlabel("Time", fontweight="bold")
plt.ylabel("Mean Score", fontweight="bold")
plt.title(" Cultural Outcomes Across Time", fontweight="bold")

plt.ylim(1, 5)
plt.legend()
plt.tight_layout()
plt.savefig('Repeated Anova',dpi=300)
plt.show()
