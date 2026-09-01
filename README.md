# vebc

**vebc** is an R package for generating metadata files for VEBC datasets.  

---

## Installation

### Option 1: Install from GitHub

You can install the package directly from GitHub using the devtools package:

```r
install.packages("devtools")

library(devtools)

devtools::install_github("miklosban/vebc")
```

### Option 2: Run locally with load_all()

If you experience problems installing the package from GitHub, for example due to GitHub authentication or token configuration, you can clone or download the repository and run the package directly from your local copy.

First, clone the repository:

git clone https://github.com/miklosban/vebc.git

Alternatively, download the repository as a ZIP file from GitHub and extract it.

Then open R or RStudio, set the working directory to the package directory, and run:

```r
library(devtools)

devtools::load_all()
```

This loads the package directly from the source code without installing it.

You can then use the package functions normally.

## Usage

```r
library(vebc)

## Generate metadata spreadsheet files for the dataset uploaded 
## into the OBM SQL table called "Behaviour" in schema "vebc"
generate_metadata_files("Behaviour", "vebc")

```

## License

This package is licensed under the GPL-3 License – see the LICENSE
file for details.

## Contributing

Contributions are welcome! Please open issues or submit pull requests on GitHub.
