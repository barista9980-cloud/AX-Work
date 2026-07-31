"""
Universal Enterprise AX Asset Management Engine - Main Entry Point
Usage:
    python main.py --setup           # Initial setup wizard: Prompts for Company Name & Drive/PC Path + Creates Folders
    python main.py --setup-dirs      # (Alias for --setup)
    python main.py --all             # Run complete asset processing pipeline
    python main.py --real-estate     # Generate Real Estate Master Register (.xlsx)
    python main.py --vehicle         # Generate Corporate Vehicle Master Register (.xlsx)
    python main.py --insurance       # Generate Corporate Insurance Master Register (.xlsx)
"""
import sys
import argparse

sys.stdout.reconfigure(encoding='utf-8')

from src.config import (
    DEFAULT_BASE_DIR,
    DEFAULT_COMPANY_NAME,
    DEFAULT_SNAPSHOT_DATE,
    run_interactive_setup,
    load_user_config
)
from src.folder_structure_engine import init_corporate_folder_structure
from src.real_estate_engine import generate_real_estate_excel
from src.vehicle_engine import generate_vehicle_excel
from src.insurance_engine import generate_insurance_excel

def main():
    parser = argparse.ArgumentParser(description="Universal Corporate AX Asset Management Framework Engine")
    parser.add_argument("--setup", action="store_true", help="Launch setup wizard to configure Company Name, Drive/PC Path, and build Folder Structure")
    parser.add_argument("--setup-dirs", action="store_true", help="Launch setup wizard and initialize corporate directory structure")
    parser.add_argument("--base-dir", default=None, help="Base directory path for asset management")
    parser.add_argument("--company", default=None, help="Corporate Legal Name")
    parser.add_argument("--snapshot-date", default=None, help="Snapshot Cut-off Date")
    parser.add_argument("--all", action="store_true", help="Execute complete asset pipeline")
    parser.add_argument("--real-estate", action="store_true", help="Generate Real Estate Master Register (.xlsx)")
    parser.add_argument("--vehicle", action="store_true", help="Generate Corporate Vehicle Master Register (.xlsx)")
    parser.add_argument("--insurance", action="store_true", help="Generate Corporate Insurance Master Register (.xlsx)")

    args = parser.parse_args()

    # If --setup or --setup-dirs is passed, run the integrated setup workflow
    if args.setup or args.setup-dirs if hasattr(args, 'setup-dirs') else args.setup_dirs:
        company, base_dir, snapshot_date = run_interactive_setup()
        init_corporate_folder_structure(base_dir)
        print("=================================================================")
        print(f"✅ [{company}] 마스터 셋팅 및 폴더 생성이 정상 완료되었습니다.")
        print(f"   • 구글드라이브/PC 경로 : {base_dir}")
        print("=================================================================
")
        return

    # Load saved user config
    cfg_company, cfg_base_dir, cfg_snapshot = load_user_config()

    target_company = args.company if args.company else cfg_company
    target_base_dir = args.base_dir if args.base_dir else cfg_base_dir
    target_snapshot = args.snapshot_date if args.snapshot_date else cfg_snapshot

    # If company name is not set yet, automatically launch setup
    if "사명 미설정" in target_company or not target_base_dir or "Enterprise_Assets" in target_base_dir:
        print("
[안내] 저장소 셋팅이 완료되지 않았습니다. 셋팅 마법사를 시작합니다.")
        target_company, target_base_dir, target_snapshot = run_interactive_setup()
        init_corporate_folder_structure(target_base_dir)
        return

    # If no specific pipeline flag passed, default to --all
    if not any([args.all, args.real_estate, args.vehicle, args.insurance]):
        args.all = True

    print("=================================================================")
    print(f"AX Enterprise Asset Management Engine ({target_company})")
    print(f"   Target Directory : {target_base_dir}")
    print(f"   Snapshot Cut-off : {target_snapshot}")
    print("=================================================================
")

    if args.all or args.real_estate:
        generate_real_estate_excel(target_base_dir, target_company, target_snapshot)

    if args.all or args.vehicle:
        generate_vehicle_excel(target_base_dir, target_company, target_snapshot)

    if args.all or args.insurance:
        generate_insurance_excel(target_base_dir, target_company, target_snapshot)

    print("=================================================================")
    print("ALL REQUESTED AX ASSET PIPELINES COMPLETED SUCCESSFULLY!")
    print("=================================================================")

if __name__ == "__main__":
    main()
