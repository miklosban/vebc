# vebc

**vebc** is an R package for generating metadata files for VEBC datasets.  

---

## Installation

You can install the package directly from GitHub using `devtools`:

## Usage

First, install and load the `devtools` package (if you haven't already):

```r
install.packages("devtools")
library(devtools)

devtools::install_github("miklosban/vebc")
library(vebc)

# Generate OBM metadata for the dataset uploaded into the table called "Behaviour" in schema "vebc"
generate_obm_metadata("Behaviour", "vebc")

```

## License

This package is licensed under the GPL-3 License – see the LICENSE
file for details.

## Contributing

Contributions are welcome! Please open issues or submit pull requests on GitHub.
