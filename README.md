
<!-- README.md is generated from README.Rmd. Please edit that file -->

# rdocx

# rdocx <img src="man/figures/hex-rdocx.png" align="right" width="80" alt="rdocx website" /></a>

<!-- badges: start -->
[![R-CMD-check](https://github.com/Novartis/rdocx/actions/workflows/check_package.yml/badge.svg)](https://github.com/Novartis/rdocx/actions/workflows/check_package.yml)
[![pkgdown](https://github.com/Novartis/rdocx/actions/workflows/pkgdown.yml/badge.svg)](https://github.com/Novartis/rdocx/actions/workflows/pkgdown.yml)
<!-- badges: end -->

## Introduction

In clinical trials we work in interdisciplinary teams where the
discussion of outputs is often facilitated using static documents. We
wanted to bring the advantages of modern tools (R, markdown, git) and
software development practices to the production of company documents.
We used an object-oriented approach to create classes for report items
with a suite of tests. Finally, the report is rendered programmatically
in docx format using a company template. This enables our statisticians
to work in a truly end to end fashion within a GxP environment with the
end product in a format suitable for interdisciplinary collaboration.

For the open-source version of our package we have replaced our company
template/s with a generic template that demonstrates the functionality
of the package. In most cases, you prefer to provide a custom template
to enforce your desired formatting. In that case, please ensure that the
classes are compatible with your template and adjust as needed.

## The use cases

As for now, there are two main applications of **{rdocx}**:

- [Generic Report](https://opensource.nibr.com/rdocx/articles/generic_report.html): As our first use case, the Generic
  Report Rmd template allows the user to create a report in RMarkdown
  and render it directly from R into a Word document following the style
  of the Generic Report Word template that is used as a reference
  document by the Generic Report Rmd template.

- [Automated Reporting](https://opensource.nibr.com/rdocx/articles/automated_reporting.html): Allows the user to automatically
  update all or selected tables/figures in a Word document.

## Installation

**{rdocx}** can be installed from Gitlab using (update before pushing to
Github)

``` r
devtools::install_github("Novartis/rdocx",
                         ref='main',
                         build_vignettes=TRUE,
                         lib='~/R')
```

Remember to set the parameter `lib` to the desired location for
installation, as it will be installed in your local directory.

## Usage

``` r
library('rdocx', lib.loc="~/R")
```

For quick access to the documentation, visit [rdocx site](https://opensource.nibr.com/rdocx/index.html). Read the vignettes on how to use the package with step-by-step
examples online:

- [Generic Report](https://opensource.nibr.com/rdocx/articles/generic_report.html)

- [Automated Reporting](https://opensource.nibr.com/rdocx/articles/automated_reporting.html)

Or call the vignettes directly in R:

``` r
vignette("generic_report", package = "rdocx")
```

``` r
vignette("automated_reporting", package = "rdocx")
```

## Class Structure Overview

The structure of the package is displayed in a simplified format below.
Note that further helper and checking functions exist that are not
visualised here for simplicity.

![](https://mermaid.ink/img/pako:eNqtlVFr2zAQx7-K0dPG2n6A0JeywR432j4axEW-2AJZ8qRTXVPy3SdbTuLIclhhDgRb99Pp_n-d7A8mTIVsx4QC535IqC20pS7C9RM1WimesTOWisf7--JVksLfUOMW8L0BXaMy9Svs1Sb1JEi-SRpuQi-y1kDeoltgT55MC4RVBKWuJzY8KRD4y1PnyRWRva6leHx4SIaeTR_JZKkJTcbObPyfzEqK_oihY6mXTKbiC5gmPPt7YsbrmyM7TrNTAk4js45W2IGlFjWtY458NWxNjEHt2z3aXDSY4Nbjb2idNDpTR9C6Ht17xyeJi9BB4TtNfovTtnC6bHXCuPN-rKAaZ1d4F6z78nXT3qQjlh5DVXFr-tPkS96R_OeUoUtyO5f3ZNPCvhm47_KT-gaIR7-qRVRjvyhyWWDa2_9D9Pps5FRraD_Zp4nka1FpEVcvkdz6MAPrSAtSc_DUGMvzVUZiTsDzW_FH3M4xxjcy3DZ5qTJ5s2VkErYBIeSVEe_8IBXmyzExRTgj1GSsNyLM1aBuZBjabHQUI6Cj0MvhbBoesKuOiq3MQSk-17BSPf7YHWvRBuOr8DGadJaMGgxLsV24rfAAXlHJSn0MaHDevAxasB1Zj3fMGl83bHcA5cJTXHL-mEXk-BeGPjNY?type=png)
