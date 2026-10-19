import pandas as pd
import json

def analyze_dataframe(df: pd.DataFrame) -> dict:
    num_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols = df.select_dtypes(include=["object","category"]).columns.tolist()

    result = {
        "shape": f"{df.shape[0]} rows × {df.shape[1]} columns",
        "columns": list(df.columns),
        "dtypes": {c: str(t) for c, t in df.dtypes.items()},
        "missing": {c: int(v) for c, v in df.isnull().sum().items() if v > 0},
        "numeric_summary": {},
        "top_values": {},
        "correlations": {},
    }

    if num_cols:
        desc = df[num_cols].describe().round(3)
        result["numeric_summary"] = desc.to_dict()
        if len(num_cols) >= 2:
            corr = df[num_cols].corr().round(3)
            # Only keep strong correlations
            for c1 in corr.columns:
                for c2 in corr.columns:
                    if c1 < c2:
                        val = corr.loc[c1, c2]
                        if abs(val) > 0.3:
                            result["correlations"][f"{c1} vs {c2}"] = round(val, 3)

    for col in cat_cols[:6]:
        vc = df[col].value_counts().head(6)
        result["top_values"][col] = vc.to_dict()

    return result
