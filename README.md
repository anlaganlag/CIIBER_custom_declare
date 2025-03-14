# Excel Converter for Declaration List

A Python application that converts Excel files for declaration purposes by combining data from multiple sources and adding fixed values.

## Process Flow

```mermaid
graph TD
    subgraph Inputs
        A[Input Excel File]
        B[Reference Excel File]
    end

    subgraph Processing
        C[Excel Converter]
        E[Copy Green Headers]
        F[Match Yellow Headers]
        G[Add Fixed Values]
    end

    subgraph Output
        H[Final Excel File]
    end

    A -->|Preserved Columns| E
    B -->|Material Code Matching| F
    E --> C
    F --> C
    G --> C
    C --> H
`
    subgraph Green_Headers
        I[NO.]
        J[DESCRIPTION]
        K[Model NO.]
        L[Qty]
        M[Unit]
        N[Unit Price]
        O[Amount]
    end

    subgraph Yellow_Headers
        P[商品编号]
        Q[申报要素]
    end

    subgraph Fixed_Values
        R[币制: 美元]
        S[原产国: 中国]
        T[目的国: 印度]
        U[货源地: 深圳特区]
        V[征免: 照章征税]
    end