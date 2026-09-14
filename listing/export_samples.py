"""Export the four client documents (proposal, invoice, change order, statement) as complete PDFs for the landing page.
Wraps the listing-image pipeline with SAMPLE_DOCS=1 so the documents keep their full print areas; run make_images.py --render afterwards
to restore the cropped renders used for the listing images."""
import os, subprocess, sys
here = os.path.dirname(os.path.abspath(__file__))
env = dict(os.environ, SAMPLE_DOCS="1")
sys.exit(subprocess.run([sys.executable, os.path.join(here, "make_images.py"), "--render"], env=env).returncode)
