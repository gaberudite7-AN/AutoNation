import xlwings as xw
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import pyodbc
import time
import warnings
import os
import glob
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")

class AllocationMapper:
    def __init__(self, base_path, automation_path):
        self.base_path = base_path
        
        # Excel files
        self.allocation_mapper_file = os.path.join(base_path, "AllocWards Override 91725.xlsx")
        self.allocation_group_file = os.path.join(automation_path, "AllocationGroup.csv")
        self.mb_sprinter_file = os.path.join(automation_path, "MB_Sprinters.csv")

    def run_NDD_sql_queries(self, queries: dict):
        try:
            with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=NDDPRDDB01.us1.autonation.com, 48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
            ) as conn:
                start_time = time.time()
                results = {}
                for name, query in queries.items():
                    df = pd.read_sql(query, conn)
                    results[name] = df
                    elapsed = time.time() - start_time
                    print(f"Loaded {name} in {elapsed:.2f} seconds")
                return results
        except Exception as e:
            print("❌ Connection failed:", e)
            return {}
    
    def read_files(self):
        
        allocation_group_df = pd.read_csv(self.allocation_group_file, dtype=str)
        mb_sprinter_df = pd.read_csv(self.mb_sprinter_file, dtype=str)
        return allocation_group_df, mb_sprinter_df

    def map_general_allocation_group(self, sales_inventory_df, allocation_group_df):

        # Handle empty or invalid allocation_group_df
        if allocation_group_df.empty or allocation_group_df.shape[1] == 0:
            sales_inventory_df["ALLOCATIONGROUP"] = "nomap"
            return sales_inventory_df
    
        # Normalize column names
        sales_inventory_df = self.normalize_columns(sales_inventory_df)
        allocation_group_df = self.normalize_columns(allocation_group_df)

        # Extract Make_Model from allocation group file
        # Example: "1282 {Acura_MDX}" -> key: "ACURA_MDX", value: "1282"
        allocation_group_df["MAKE_MODEL"] = allocation_group_df.iloc[:,0].str.extract(r"\{(.+?)\}")[0].str.upper()
        allocation_group_df["CODE"] = allocation_group_df.iloc[:,0].str.extract(r"^(\d+)")[0]
        allocation_map = dict(zip(allocation_group_df["MAKE_MODEL"], allocation_group_df.iloc[:,0]))

        # Abbreviation logic for models
        def abbreviate_model(row):
            model = row["MODEL"].strip().upper()
            make = row["MAKE"].strip().upper()

            # Add more abbreviations as needed
            if make == "CADILLAC" and "ESCALADE" in model:
                model = "ESC"
            elif make == "AUDI" and "SQ5" in model:
                model = "Q5"
            elif make == "AUDI" and "RS Q8" in model:
                model = "RS"
            elif make == "AUDI" and "SQ7" in model:
                model = "Q7"
            elif make == "AUDI" and "SQ8" in model:
                model = "Q8"
            elif make == "AUDI" and "RS e-tron GT" in model:
                model = "ETRON GT"
            elif make == "HYUNDAI" and "PALISADE" in model:
                model = "PALISADE"
            elif make == "FORD" and "TRANSIT" in model:
                model = "TRANSIT"            
            elif make == "FORD" and "E-TRANSIT" in model:
                model = "E-TRANSIT"
            elif make == "TOYOTA" and "PRIUS" in model:
                model = "PRIUS"
            elif make == "CHEVROLET" and "CORVETTE" in model:
                model = "CORVET"
            elif make == "JEEP" and "GRAND WAGONEER" in model:
                model = "GRAND WAGONEER"
            # May need to adjust this??
            elif make == "JEEP" and model == "GRAND":
                model = "GRAND WAGONEER"   
            elif make == "JEEP" and "WAGONEER" in model:
                if 'SPORT' in model:
                    model = "WAGONEER S"
                else:
                    model = "WAGONEER"
            elif make == "GMC" and "YUKON" in model:
                if "XL" in model:
                    model = "YKNXL"
                else:
                    model = "YKN"
            elif make == "MAZDA" and "HATCHBACK" in model:
                model = "MAZDA3 HB"
            return f"{make}_{model}"
        
        # Run function to create abbreviated model names
        sales_inventory_df["MAKE_MODEL"] = sales_inventory_df.apply(abbreviate_model, axis=1)

        # Normalize data
        # allocation_group_df["MAKE_MODEL"] = allocation_group_df.iloc[:,0].str.extract(r"\{(.+?)\}")[0].str.strip().str.upper()
        # sales_inventory_df["MAKE_MODEL"] = sales_inventory_df["MAKE_MODEL"].str.strip().str.upper()

        # Map AllocationGroup
        sales_inventory_df["ALLOCATIONGROUP"] = sales_inventory_df["MAKE_MODEL"].map(allocation_map)
        sales_inventory_df["ALLOCATIONGROUP"] = sales_inventory_df["ALLOCATIONGROUP"].fillna("nomap")
        return sales_inventory_df
    
    def map_lexus_allocation_group(self, sales_inventory_df, allocation_group_df):
        
        # Normalize column names
        sales_inventory_df = self.normalize_columns(sales_inventory_df)
        allocation_group_df = self.normalize_columns(allocation_group_df)

        # Extract Make_Model_Style from allocation group file
        allocation_group_df["MAKE_MODEL_STYLE"] = allocation_group_df.iloc[:,0].str.extract(r"\{(.+?)\}")[0].str.upper()
        allocation_map = dict(zip(allocation_group_df["MAKE_MODEL_STYLE"], allocation_group_df.iloc[:,0]))

        def build_lexus_key(row):
            make = row["MAKE"].strip().upper()
            model = row["MODEL"].strip().upper()
            style = row["STYLENAME"].strip().upper()
            if make == "LEXUS":
                style_main = " ".join(style.split()[:2])# For Lexus, use Model and StyleName together (Styename contains model already)
                return f"{make}_{style_main}"
            else:
                return None

        # Apply only to Lexus rows
        lexus_mask = sales_inventory_df["MAKE"].str.upper() == "LEXUS"
        sales_inventory_df.loc[lexus_mask, "MAKE_MODEL_STYLE"] = sales_inventory_df[lexus_mask].apply(build_lexus_key, axis=1)
        sales_inventory_df.loc[lexus_mask, "ALLOCATIONGROUP"] = sales_inventory_df.loc[lexus_mask, "MAKE_MODEL_STYLE"].map(allocation_map)
        sales_inventory_df.loc[lexus_mask, "ALLOCATIONGROUP"] = sales_inventory_df.loc[lexus_mask, "ALLOCATIONGROUP"].fillna("nomap")
        
        return sales_inventory_df
    
    def map_complications(self, sales_inventory_df):
        
        # Normalize column names
        sales_inventory_df = self.normalize_columns(sales_inventory_df)        
        
        def key(row):
            make = row["MAKE"].strip().upper()
            model = row["MODEL"].strip().upper()
            style = row["STYLENAME"].strip().upper()
            trim = row["TRIM"].strip().upper()

            # Define model groups (all uppercase for matching)
            silverado_models = (
                "SILVERADO 1500", "SILVERADO 2500", "SILVERADO 2500HD",
                "SILVERADO 3500", "SILVERADO 3500HD", "SILVERADO 3500HD CC"
            )
            ford_models = (
                "SUPER DUTY F-250 SRW", "SUPER DUTY F-250 DRW", "SUPER DUTY F-350 SRW", "SUPER DUTY F-350 DRW", "SUPER DUTY F-450 SRW", "SUPER DUTY F-450 DRW", 
                "SUPER DUTY F-550 SRW", "SUPER DUTY F-550 DRW", "SUPER DUTY F-600 SRW", "SUPER DUTY F-600 DRW"
            )
            hummer_models = (
                "HUMMER EV PICKUP", "HUMMER EV SUV"
            )
            gmc_models = (
                "SIERRA 1500", "SIERRA 2500", "SIERRA 2500HD",
                "SIERRA 3500", "SIERRA 3500HD", "SIERRA 4500HD"
            )
            ram_models = (
                "1500", "2500", "3500", "4500 Chassis Cab"
            )

            # Mapping logic for complicated cases
            # Chevrolet Silverado Crew Cab
            if make == "CHEVROLET" and model in silverado_models and "CREW CAB" in style:
                if "HD" in model:
                    return "1466 {Chevrolet_CHDCRW}"
                else:
                    return "1475 {Chevrolet_CLDCRW}"
            # Silverado Double Cab
            elif make == "CHEVROLET" and model in silverado_models and "DOUBLE CAB" in style:
                if "HD" in model:
                    return "2558 {Chevrolet_CHDDBL}"
                else: 
                    return "2558 {Chevrolet_CLDDBL}"
            # Silverado Reg
            elif make == "CHEVROLET" and model in silverado_models and "REG" in style:
                if "HD" in model:
                    return "1467 {Chevrolet_CHDREG}"
                else:
                    return "1477 {Chevrolet_CLDREG}"
            # GMC Crew Cab
            elif make == "GMC" and model in gmc_models and "CREW CAB" in style:
                if "HD" in model:
                    return "1761 {GMC_GHDCRW}"
                else:
                    return "1763 {GMC_GLDCRW}"
            # GMC Double Cab
            elif make == "GMC" and model in gmc_models and "DOUBLE CAB" in style:
                if "HD" in model:
                    return "2566 {GMC_GHDDBL}"
                else:
                    return "1764 {GMC_GLDDBL}"
            # GMC Reg
            elif make == "GMC" and model in gmc_models and "REG" in style:
                if "HD" in model:
                    return "2602 {GMC_GHDREG}"
                else: 
                    return "1765 {GMC_GLDREG}"
            # Ford Super Duty
            elif make == "FORD" and model in ford_models:
                return "1706 {Ford_F-SERIES SD}"
            # Hummer
            elif make == "GMC" and model in hummer_models:
                return "2741 {GMC_HEVTRK}"
            # Chevrolet Suburban
            elif make == "CHEVROLET" and model == "SUBURBAN":
                return "1517 {Chevrolet_SUBURB}"
            # Chevrolet Tahoe
            elif make == "CHEVROLET" and model == "TAHOE":
                return "1288 {Chevrolet_TAHOE}"
            # Chevrolet Trax
            elif make == "CHEVROLET" and model == "TRAX":
                return "1288 {Chevrolet_TRAX}"
            # Chevrolet Blazer
            elif make == "CHEVROLET" and model == "BLAZER EV":
                return "2776 {Chevrolet_BLAZEV}"
            # Toyota Bz4x
            if make == "TOYOTA" and model == "BZ":
                return "2736 {Toyota_BZ4X}"
            # Toyota Tundra
            elif make == "TOYOTA" and model == "TUNDRA":
                if "4WD" in style:
                    return "2487 {Toyota_TUNDRA 4WD}"
                else:
                    return "2488 {Toyota_TUNDRA 2WD}"
            # Toyota Corolla
            elif make == "TOYOTA" and model == "COROLLA":
                if "HYBRID" in style:
                    return "2453 {Toyota_COROLLA HYBRID}"
                else:
                    return "2451 {Toyota_COROLLA}"
            elif make == "RAM":
                if "1500" in model:
                    return "2361 {Ram_CREWPRM}"
                elif "2500" in model:
                    return "2364 {RAM_HD 2500}"
                elif "3500" in model:
                    return "2364 {RAM_HD 3500}"
                elif "4500 CHASSIS CAB" in model:
                    return "2349 {RAM_4500/5500}"
                elif "PROMASTER" in model:
                    if "EV" in model:
                        return "2817 {Ram_PROMASTER EV}"
                    if "CITY" in model:
                        return "2371 {Ram_PROMASTER CITY}"
                    else:
                        return "2368 {RAM_PROMASTER}"
                return None
        # Apply only to relevant makes
        relevant_makes = ["CHEVROLET", "GMC", "FORD", "RAM", "TOYOTA"]
        general_mask = sales_inventory_df["MAKE"].str.upper().isin(relevant_makes)
        # Only update AllocationGroup if key(row) returns a value
        new_alloc = sales_inventory_df[general_mask].apply(key, axis=1)
        update_mask = general_mask & new_alloc.notnull()
        sales_inventory_df.loc[update_mask, "ALLOCATIONGROUP"] = new_alloc[update_mask]
        return sales_inventory_df
    
    def map_complications_pipeline(self, pipeline_df):
        
        # Normalize column names
        pipeline_df = self.normalize_columns(pipeline_df)        
        
        def key(row):
            make = row["MAKE"].strip().upper()
            model = row["MODEL"].strip().upper()
            trim = row["TRIM"].strip().upper()

            # Mapping logic for complicated cases
            # Toyota Bz4x
            if make == "TOYOTA" and model == "BZ":
                return "2736 {Toyota_BZ4X}"
            # Toyota Tundra
            elif make == "TOYOTA" and "TUNDRA" in model:
                if "4-DOOR" in trim:
                    return "2487 {Toyota_TUNDRA 4WD}"
                else:
                    return "2488 {Toyota_TUNDRA 2WD}"        
            return None
        
        # Apply only to relevant makes
        relevant_makes = ["CHEVROLET", "GMC", "FORD", "RAM", "TOYOTA"]
        general_mask = pipeline_df["MAKE"].str.upper().isin(relevant_makes)
        # Only update AllocationGroup if key(row) returns a value
        new_alloc = pipeline_df[general_mask].apply(key, axis=1)
        update_mask = general_mask & new_alloc.notnull()
        pipeline_df.loc[update_mask, "ALLOCATIONGROUP"] = new_alloc[update_mask]
        return pipeline_df
    
    def map_mercedes(self, sales_inventory_df, allocation_group_df):
        # Extract Make_ModelID from allocation group file
        allocation_group_df["MAKE_MODEL_ID"] = allocation_group_df.iloc[:,0].str.extract(r"\{(.+?)\}")[0].str.upper()
        allocation_map = dict(zip(allocation_group_df["MAKE_MODEL_ID"], allocation_group_df.iloc[:,0]))

        def abbreviate_model(row):
            model_id = row["MODEL_ID"].strip().upper()
            make = row["MAKE"].strip().upper()
            return f"{make}_{model_id}"

        mercedes_mask = sales_inventory_df["MAKE"].str.upper() == "MERCEDES-BENZ"
        sales_inventory_df.loc[mercedes_mask, "MAKE_MODEL_ID"] = sales_inventory_df[mercedes_mask].apply(abbreviate_model, axis=1)
        sales_inventory_df.loc[mercedes_mask, "ALLOCATIONGROUP"] = sales_inventory_df.loc[mercedes_mask, "MAKE_MODEL_ID"].map(allocation_map)
        sales_inventory_df.loc[mercedes_mask, "ALLOCATIONGROUP"] = sales_inventory_df.loc[mercedes_mask, "ALLOCATIONGROUP"].fillna("nomap")
        return sales_inventory_df
    
    def map_mb_sprinter(self, sales_inventory_df, mb_sprinter_df):
        
        # Normalize column names
        sales_inventory_df = self.normalize_columns(sales_inventory_df)
        mb_sprinter_df = self.normalize_columns(mb_sprinter_df)
        
        # Build mapping from Model_ID to AllocationGroup
        mb_sprinter_df["MODEL_ID"] = mb_sprinter_df["MODEL_ID"].str.strip().str.upper()
        mb_sprinter_df["ALLOCATIONGROUP"] = mb_sprinter_df["ALLOCATIONGROUP"].str.strip()
        sprinter_map = dict(zip(mb_sprinter_df["MODEL_ID"], mb_sprinter_df["ALLOCATIONGROUP"]))

        # Apply mapping for Mercedes-Benz Sprinter vehicles
        sprinter_mask = (
            (sales_inventory_df["MAKE"].str.upper() == "MERCEDES-BENZ") &
            (sales_inventory_df["MODEL"].str.upper().str.contains("SPRINTER"))
        )

        sales_inventory_df.loc[sprinter_mask, "MODEL_ID"] = sales_inventory_df.loc[sprinter_mask, "MODEL_ID"].str.strip().str.upper()
        sales_inventory_df.loc[sprinter_mask, "ALLOCATIONGROUP"] = sales_inventory_df.loc[sprinter_mask, "MODEL_ID"].map(sprinter_map)
        sales_inventory_df.loc[sprinter_mask, "ALLOCATIONGROUP"] = sales_inventory_df.loc[sprinter_mask, "ALLOCATIONGROUP"].fillna("nomap")
        return sales_inventory_df

    def map_jeep(self, df, allocation_group_df):
        # Normalize columns
        df = self.normalize_columns(df)
        allocation_group_df = self.normalize_columns(allocation_group_df)

        # Extract Make_ModelID from allocation group file
        allocation_group_df["MAKE_MODEL_ID"] = allocation_group_df.iloc[:,0].str.extract(r"\{(.+?)\}")[0].str.upper()
        allocation_map = dict(zip(allocation_group_df["MAKE_MODEL_ID"], allocation_group_df.iloc[:,0]))

        # Model code mapping
        model_codes = {
            "JLJL74": "SPORT", "JLJP74": "SAHARA", "JLJS84": "RUBICON", "JLJX74": "RUBICON 392",
            "JLXL74": "SPORT", "JLXP74": "SAHARA", "JLXS74": "RUBICON", "JLJL72": "2DR", "JLJS72": "2DR"
        }

        # Model abbreviation mapping
        model_abbr = {
            "WRANGLER": "WRNGL",
            "CHEROKEE": "CHRK",
            "GRAND CHEROKEE": "GRCHRK"
            # Add more as needed
        }

        def abbreviate_model(row):
            make = row["MAKE"].strip().upper()
            model = row["MODEL"].strip().upper()
            model_id = row.get("MODEL_ID", "").strip().upper()
            abbr = model_abbr.get(model, model)
            if make == "JEEP" and model_id in model_codes:
                code = model_codes[model_id]
                return f"{make}_{abbr} {code}"
            return f"{make}_{model}"

        # Only update Jeep rows that are still 'nomap' or empty
        jeep_mask = (
            (df["MAKE"].str.upper() == "JEEP") &
            ((df["ALLOCATIONGROUP"].isna()) | (df["ALLOCATIONGROUP"].str.lower() == "nomap"))
        )
        df.loc[jeep_mask, "MAKE_MODEL_ID"] = df[jeep_mask].apply(abbreviate_model, axis=1)
        df.loc[jeep_mask, "ALLOCATIONGROUP"] = df.loc[jeep_mask, "MAKE_MODEL_ID"].map(allocation_map)
        df.loc[jeep_mask, "ALLOCATIONGROUP"] = df.loc[jeep_mask, "ALLOCATIONGROUP"].fillna("nomap")
        return df


    def fill_nomap_allocationgroups(self, df, run_query_func, allocation_group_df):
        """
        For rows where ALLOCATIONGROUP is 'nomap', query the DB for MODEL_IDs and fill in ALLOCATIONGROUP if found.
        Flag these rows, then rebuild MAKE_MODEL (using ALLOCATIONGROUP), and run them through map_general_allocation_group.
        """
        # Flag rows that were nomap
        df["WAS_NOMAP"] = df["ALLOCATIONGROUP"].str.lower() == "nomap"
        nomap_mask = df["WAS_NOMAP"]
        missing_model_ids = df.loc[nomap_mask, "MODEL_ID"].unique().tolist()
        if not missing_model_ids:
            return df

        # Build dynamic query
        model_id_str = "', '".join(missing_model_ids)
        dynamic_query = f"""
        WITH RankedInventory AS (
            SELECT 
                accountingmonth,
                year,
                Model_ID,
                Make,
                model,
                trim,
                AllocationGroup,
                ROW_NUMBER() OVER (
                    PARTITION BY Model_ID 
                    ORDER BY year DESC, accountingmonth DESC
                ) AS rn
            FROM [NDD_ADP_RAW].[NDDUsers].[vInventoryMonthEnd]
            WHERE Model_ID IN ('{model_id_str}')
        )
        SELECT 
            accountingmonth,
            year,
            Model_ID,
            Make,
            model,
            trim,
            AllocationGroup
        FROM RankedInventory
        WHERE rn = 1
        """

        # Run the query
        search_results = run_query_func({"Search": dynamic_query})
        search_df = search_results.get("Search")
        if search_df is not None and not search_df.empty:
            search_df = self.normalize_columns(search_df)
            # Map MODEL_ID to allocation group
            modelid_to_alloc = dict(zip(search_df["MODEL_ID"], search_df["ALLOCATIONGROUP"]))

            # Update ALLOCATIONGROUP for flagged rows
            df.loc[nomap_mask, "ALLOCATIONGROUP"] = (
                df.loc[nomap_mask, "MODEL_ID"].map(modelid_to_alloc).fillna("nomap")
            )

            # Rebuild MAKE_MODEL for flagged rows using the new ALLOCATIONGROUP value
            df.loc[nomap_mask, "MAKE_MODEL"] = (
                df.loc[nomap_mask, "MAKE"].str.strip().str.upper() + "_" +
                df.loc[nomap_mask, "ALLOCATIONGROUP"].str.strip().str.upper()
            )

            # Run only the flagged rows through map_general_allocation_group
            flagged_rows = df.loc[nomap_mask].copy()
            flagged_rows = self.map_general_allocation_group(flagged_rows, allocation_group_df)
            df.loc[nomap_mask, "ALLOCATIONGROUP"] = flagged_rows["ALLOCATIONGROUP"]

        return df

    def delete_unneeded_makes(self, sales_inventory_df):
        makes_to_delete = ['Rolls-Royce', 'Maserati', 'McLaren', 'Mitsubishi', 'INEOS']
        models_to_delete = ['Police Interceptor Utility', 'MEDIUM TRK']
        mask = (~sales_inventory_df['MODEL'].isin(models_to_delete)) & (~sales_inventory_df['MAKE'].isin(makes_to_delete))
        return sales_inventory_df[mask]

    def normalize_columns(self, df):
        df.columns = [col.strip().upper() for col in df.columns]
        return df

if __name__ == "__main__":
    
    # Configure path 
    base_path = r"W:\Corporate\Inventory\Autos\Alloc Grp & Wards Map\2025"
    automation_path = r"W:\Corporate\Inventory\Autos\Alloc Grp & Wards Map\2025\Automation"
    AllocationMapper = AllocationMapper(base_path, automation_path)

    # Update the Sales_Inventory data
    Sales_Inventory_query = """
    SELECT vInventoryMonthEnd.StyleID, 
    vInventoryMonthEnd.Year, 
    vInventoryMonthEnd.Make, 
    vInventoryMonthEnd.Model, 
    vInventoryMonthEnd.Model_ID, 
    vInventoryMonthEnd.StyleName, 
    vInventoryMonthEnd.Trim

    FROM NDD_ADP_RAW.NDDUsers.vInventoryMonthEnd vInventoryMonthEnd

    WHERE vInventoryMonthEnd.AllocationGroup='nomap'
    AND vInventoryMonthEnd.Department='300'
    AND vInventoryMonthEnd.AccountingMonth>'12/1/2020'
    AND vInventoryMonthEnd.Year > 2022 
    AND vInventoryMonthEnd.Make not in ('Kia','Rivian','Tesla','')
    AND vInventoryMonthEnd.Model_ID is not NULL
    AND vInventoryMonthEnd.Model not in ('F-150 Police Responder','F-59 Commercial Stripped Chassis')

    GROUP BY vInventoryMonthEnd.StyleID, 
    vInventoryMonthEnd.Year, 
    vInventoryMonthEnd.Make, 
    vInventoryMonthEnd.Model, 
    vInventoryMonthEnd.Model_ID, 
    vInventoryMonthEnd.StyleName, 
    vInventoryMonthEnd.Trim

    ORDER BY 3
    """
    AllocationGroupPipeline_query = """
    SELECT modelyear, 
    make, 
    model, 
    model_number, 
    alloc_grp_map, 
    Trim

    FROM NDDUsers.vOnOrderInfo_Dly_SnapShot

    WHERE update_date>'2021-02-27'
    and modelyear > 2022
    and model <> 'Motor Home'

    GROUP BY modelyear, 
    make, 
    model, 
    model_number, 
    alloc_grp_map, 
    Trim

    HAVING (modelyear>=2019 and make<>'nomake' and alloc_grp_map='nomap' or alloc_grp_map='' or alloc_grp_map is null)

    ORDER BY make, 
    model,
    model_number;"""
    Search_query = """
WITH RankedInventory AS (
    SELECT 
        accountingmonth,
        year,
        Model_ID,
        Make,
        model,
        trim,
        AllocationGroup,
        ROW_NUMBER() OVER (
            PARTITION BY Model_ID 
            ORDER BY year DESC, accountingmonth DESC
        ) AS rn
    FROM [NDD_ADP_RAW].[NDDUsers].[vInventoryMonthEnd]
    WHERE Model_ID IN ('8346', 'MPJM74', 'DT6M98')
)

SELECT 
    accountingmonth,
    year,
    Model_ID,
    Make,
    model,
    trim,
    AllocationGroup
FROM RankedInventory
WHERE rn = 1
"""
    
    queries = {
        "Sales_Inventory": Sales_Inventory_query,
        "AllocationGroupPipeline": AllocationGroupPipeline_query
    }

    query_results = AllocationMapper.run_NDD_sql_queries(queries)
    allocation_group_df, mb_sprinter_df = AllocationMapper.read_files()
    # Rename columns
    query_results["AllocationGroupPipeline"] = query_results["AllocationGroupPipeline"].rename(columns={'model_number': 'MODEL_ID'})
    
    # fillnas
    query_results["AllocationGroupPipeline"]["MODEL_ID"] = query_results["AllocationGroupPipeline"]["MODEL_ID"].fillna(0).astype(str)

    # # Apply separate mapping methods for each Brand
    query_results["Sales_Inventory"] = AllocationMapper.map_general_allocation_group(query_results["Sales_Inventory"], allocation_group_df)
    query_results["Sales_Inventory"] = AllocationMapper.map_lexus_allocation_group(query_results["Sales_Inventory"], allocation_group_df)
    query_results["Sales_Inventory"] = AllocationMapper.map_complications(query_results["Sales_Inventory"])
    query_results["Sales_Inventory"] = AllocationMapper.map_mercedes(query_results["Sales_Inventory"], allocation_group_df)
    query_results["Sales_Inventory"] = AllocationMapper.map_mb_sprinter(query_results["Sales_Inventory"], mb_sprinter_df)
    query_results["Sales_Inventory"] = AllocationMapper.map_jeep(query_results["Sales_Inventory"], allocation_group_df)
    query_results["Sales_Inventory"] = AllocationMapper.delete_unneeded_makes(query_results["Sales_Inventory"])

    # Output to CSV
    query_results["Sales_Inventory"].to_csv(os.path.join(automation_path, "Sales_Inventory.csv"), index=False)
    print(f"✅ Sales_Inventory mapping complete. {len(query_results['Sales_Inventory'])} rows written to CSV.")

    # Apply general mapping
    query_results["AllocationGroupPipeline"] = AllocationMapper.map_general_allocation_group(query_results["AllocationGroupPipeline"], allocation_group_df)    
    query_results["AllocationGroupPipeline"] = AllocationMapper.map_mercedes(query_results["AllocationGroupPipeline"], allocation_group_df)
    query_results["AllocationGroupPipeline"] = AllocationMapper.map_mb_sprinter(query_results["AllocationGroupPipeline"], mb_sprinter_df)
    query_results["AllocationGroupPipeline"] = AllocationMapper.map_jeep(query_results["AllocationGroupPipeline"], allocation_group_df)
    query_results["AllocationGroupPipeline"] = AllocationMapper.delete_unneeded_makes(query_results["AllocationGroupPipeline"])
   
    # Apply lookup query for nomap rows
    query_results["AllocationGroupPipeline"] = AllocationMapper.fill_nomap_allocationgroups(
        query_results["AllocationGroupPipeline"], AllocationMapper.run_NDD_sql_queries, allocation_group_df
    )

    # Output to CSV
    query_results["AllocationGroupPipeline"].to_csv(os.path.join(automation_path, "Pipeline.csv"), index=False)
    print(f"✅ AllocationGroupPipeline mapping complete. {len(query_results['AllocationGroupPipeline'])} rows written to CSV.")


    # Delete medium trk
    # Jeeps - need to lookup (wrangler and gladiator)
    # 1517 for suburb
    # tahoe and promaster