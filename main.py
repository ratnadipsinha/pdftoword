"""
PDF2Word Pro — Entry Point
Supports:
  GUI mode  : python main.py
  CLI mode  : python main.py input.pdf [output.docx]
  Batch mode: python main.py --batch input_folder/ output_folder/
"""

import sys
import os


def run_gui():
    from src.gui import PDF2WordApp
    app = PDF2WordApp()
    app.run()


def run_cli():
    from src.converter import PDF2WordConverter

    args = sys.argv[1:]

    # Batch mode: --batch <input_folder> <output_folder>
    if "--batch" in args:
        idx = args.index("--batch")
        try:
            input_dir = args[idx + 1]
            output_dir = args[idx + 2]
        except IndexError:
            print("Usage: python main.py --batch <input_folder> <output_folder>")
            sys.exit(1)

        def progress(current, total, name):
            print(f"  [{current}/{total}] {name}")

        converter = PDF2WordConverter(progress_cb=progress)
        print(f"\nConverting PDFs in: {input_dir}")
        results = converter.convert_batch(input_dir, output_dir)
        PDF2WordConverter.print_summary(results)
        failed = sum(1 for r in results if not r.success)
        sys.exit(failed)

    # Single file: python main.py input.pdf [output.docx]
    if len(args) >= 1:
        input_pdf = args[0]
        output_docx = args[1] if len(args) >= 2 else None

        converter = PDF2WordConverter()
        print(f"\nConverting: {input_pdf}")
        result = converter.convert(input_pdf, output_docx)

        if result.success:
            print(f"Done: {result.output_path}")
            print(f"Pages: {result.page_count} | Time: {result.duration:.1f}s")
        else:
            print(f"FAILED: {result.error}", file=sys.stderr)
            sys.exit(1)
        return

    # No args → launch GUI
    run_gui()


if __name__ == "__main__":
    if len(sys.argv) == 1:
        run_gui()
    else:
        run_cli()
