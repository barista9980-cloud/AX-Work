"""
Universal Enterprise AX Asset Management Engine - Main Entry Point
Usage:
    python main.py --all
    python main.py --real-estate
    python main.py --vehicle
    python main.py --insurance
    python main.py --setup-dirs
"""
import sys
import argparse

sys.stdout.reconfigure(encoding='utf-8')

from src.config import DEFAULT_BASE_DIR, DEFAULT_COMPANY_NAME, DEFAULT_SNAPSHOT_DATE
from src.folder_structure_engine import init_corporate_folder_structure
from src.real_estate_engine import generate_real_estate_excel
from src.vehicle_engine import generate_vehicle_excel
from src.insurance_engine import generate_insurance_excel

def main():
    parser = argparse.ArgumentParser(description="Universal Corporate AX Asset Management Framework Engine")
    parser.add_argument("--base-dir", default=DEFAULT_BASE_DIR, help="Base directory path for asset management")
    parser.add_argument("--company", default=DEFAULT_COMPANY_NAME, help="Corporate Legal Name")
    parser.add_argument("--snapshot-date", default=DEFAULT_SNAPSHOT_DATE, help="Snapshot Cut-off Date")
    parser.add_argument("--all", action="store_true", help="Execute complete asset pipeline")
    parser.add_argument("--setup-dirs", action="store_true", help="Initialize corporate directory structure and upload queues")
    parser.add_argument("--real-estate", action="store_true", help="Generate Real Estate Master Register (.xlsx)")
    parser.add_argument("--vehicle", action="store_true", help="Generate Corporate Vehicle Master Register (.xlsx)")
    parser.add_argument("--insurance", action="store_true", help="Generate Corporate Insurance Master Register (.xlsx)")

    args = parser.parse_args()

    # If no flags passed, default to --all
    if not any([args.all, args.setup_dirs, args.real_estate, args.vehicle, args.insurance]):
        args.all = True

    print("=================================================================")
    print(f"AX Enterprise Asset Management Engine ({args.company})")
    print(f"   Target Directory : {args.base_dir}")
    print(f"   Snapshot Cut-off : {args.snapshot_date}")
    print("=================================================================\n")

    if args.setup_dirs or args.all:
        init_corporate_folder_structure(args.base_dir)

    if args.real_estate or args.all:
        generate_real_estate_excel(args.base_dir, args.company, args.snapshot_date)

    if args.vehicle or args.all:
        generate_vehicle_excel(args.base_dir, args.company, args.snapshot_date)

    if args.insurance or args.all:
        generate_insurance_excel(args.base_dir, args.company, args.snapshot_date)

    print("=================================================================")
    print("ALL REQUESTED AX ASSET PIPELINES COMPLETED SUCCESSFULLY!")
    print("=================================================================")

if __name__ == "__main__":
    main()
