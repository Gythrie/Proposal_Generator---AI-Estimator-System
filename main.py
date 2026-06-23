import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "modules"))

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Team 3 Proposal Generation Pipeline")
    parser.add_argument("--team1", default="data/dummy_team1_output.json")
    parser.add_argument("--team2", default="data/dummy_team2_output.json")
    args = parser.parse_args()

    from proposal_engine import run_track2, run_track3

    print("\n" + "=" * 60)
    print("PROPOSAL GENERATION PIPELINE")
    print("=" * 60 + "\n")

    output = run_track2(team1_path=args.team1, team2_path=args.team2)
    print(f"\n  Proposal saved → {output}\n")

    report_path = run_track3(team2_path=args.team2)
    print(f"\n  Validated proposal → {report_path}\n")

    print("=" * 60)
    print("  PIPELINE COMPLETE")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    main()