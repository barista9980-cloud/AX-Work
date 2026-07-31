"""
Universal Enterprise AX Asset Management Engine - Main Entry Point
Usage:
    python main.py --setup           # Initial setup wizard for Company Name & Drive/PC Path
    python main.py --all             # Run full pipeline with saved config
    python main.py --real-estate     # Run real estate pipeline
    python main.py --vehicle         # Run vehicle pipeline
    python main.py --insurance       # Run insurance pipeline
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
    parser.add_argument("--setup", action="store_true", help="Launch interactive wizard to setup Company Name and Drive/PC Path")
    parser.add_argument("--base-dir", default=None, help="Base directory path for asset management")
    parser.add_argument("--company", default=None, help="Corporate Legal Name")
    parser.add_argument("--snapshot-date", default=None, help="Snapshot Cut-off Date")
    parser.add_argument("--all", action="store_true", help="Execute complete asset pipeline")
    parser.add_argument("--setup-dirs", action="store_true", help="Initialize corporate directory structure and upload queues")
    parser.add_argument("--real-estate", action="store_true", help="Generate Real Estate Master Register (.xlsx)")
    parser.add_argument("--vehicle", action="store_true", help="Generate Corporate Vehicle Master Register (.xlsx)")
    parser.add_argument("--insurance", action="store_true", help="Generate Corporate Insurance Master Register (.xlsx)")

    args = parser.parse_args()

    # If --setup flag is provided or no config exists, run interactive wizard
    if args.setup:
        company, base_dir, snapshot_date = run_interactive_setup()
        return

    # Load configured defaults
    cfg_company, cfg_base_dir, cfg_snapshot = load_user_config()

    target_company = args.company if args.company else cfg_company
    target_base_dir = args.base_dir if args.base_dir else cfg_base_dir
    target_snapshot = args.snapshot_date if args.snapshot_date else cfg_snapshot

    # If company name is not configured yet, suggest interactive setup
    if "사명 미설정" in target_company:
        print("
[안내] 아직 기업 사명 및 경로가 설정되지 않았습니다.")
        target_company, target_base_dir, target_snapshot = run_interactive_setup()

    # If no specific execution flags passed, default to --all
    if not any([args.all, args.setup_dirs, args.real_estate, args.vehicle, args.insurance]):
        args.all = True

    print("=================================================================")
    print(f"AX Enterprise Asset Management Engine ({target_company})")
    print(f"   Target Directory : {target_base_dir}")
    print(f"   Snapshot Cut-off : {target_snapshot}")
    print("=================================================================
")

    if args.setup_dirs or args.all:
        init_corporate_folder_structure(target_base_dir)

    if args.real_estate or args.all:
        generate_real_estate_excel(target_base_dir, target_company, target_snapshot)

    if args.vehicle or args.all:
        generate_vehicle_excel(target_base_dir, target_company, target_snapshot)

    if args.insurance or args.all:
        generate_insurance_excel(target_base_dir, target_company, target_snapshot)

    print("=================================================================")
    print("ALL REQUESTED AX ASSET PIPELINES COMPLETED SUCCESSFULLY!")
    print("=================================================================")

if __name__ == "__main__":
    main()
