# School meals pilot: the attendance model
library(dplyr)
library(fixest)

attendance <- read.csv("attendance.csv")

# The sample: schools observed in every year, 2019 to 2025
panel <- attendance |>
  group_by(school) |>
  filter(n_distinct(year) == 7) |>
  ungroup()

# Main model: difference-in-differences with school and year fixed effects
main <- feols(attendance_rate ~ pilot:post | school + year,
              data = panel, cluster = ~school)
summary(main)

# Robustness: without the first pilot year
robust <- feols(attendance_rate ~ pilot:post | school + year,
                data = filter(panel, year != 2023), cluster = ~school)
summary(robust)
