import os
import unittest
import zipfile
import HERMES_Logo_change_V1 as hermes
from unittest.mock import MagicMock, patch

class MockIntVar:
    """A mock class for tk.IntVar to avoid initializing tkinter."""
    def set(self, value):
        pass

class TestBug(unittest.TestCase):

    def setUp(self):
        # Create a temporary directory for our test files
        self.test_dir = "temp_test_dir"
        os.makedirs(self.test_dir, exist_ok=True)

        # Create a dummy logo file
        self.logo_path = os.path.join(self.test_dir, "logo.png")
        with open(self.logo_path, "w") as f:
            f.write("this is the logo")

        # Create the file that is at risk of being overwritten
        self.image1_path = os.path.join(self.test_dir, "image1.png")
        self.image1_original_content = "this is important data"
        with open(self.image1_path, "w") as f:
            f.write(self.image1_original_content)

        # Create a dummy dotx file (which is a zip file)
        self.vorlagen_dir = os.path.join(self.test_dir, "vorlagen")
        os.makedirs(self.vorlagen_dir, exist_ok=True)
        dotx_path = os.path.join(self.vorlagen_dir, "template.dotx")
        with zipfile.ZipFile(dotx_path, 'w') as zf:
            zf.writestr("word/media/image1.png", "dummy image data")

        # Mock the GUI elements that the thread function interacts with
        self.mock_fortschritt_var = MockIntVar()
        self.mock_status_label = MagicMock()
        self.mock_ergebnisse_text = MagicMock()
        self.mock_ergebnisse_button = MagicMock()

        # Mock the global GUI elements that the thread function interacts with
        hermes.start_button = MagicMock()
        hermes.progress_bar = MagicMock()
        hermes.status_label = MagicMock()

        # Mock the 'fenster' object and its necessary methods/attributes
        mock_fenster = MagicMock()
        mock_fenster.update_idletasks = MagicMock()
        mock_fenster.ergebnis_fenster.deiconify = MagicMock()
        hermes.fenster = mock_fenster

    def tearDown(self):
        # Clean up the temporary directory and files
        import shutil
        shutil.rmtree(self.test_dir)
        # Unset the mocked global
        if hasattr(hermes, 'fenster'):
            del hermes.fenster

    @patch('HERMES_Logo_change_V1.messagebox')
    def test_overwrite_bug_is_fixed(self, mock_messagebox):
        # Run the file processing function
        hermes.dateien_bearbeiten_thread(
            self.vorlagen_dir,
            self.logo_path,
            self.mock_fortschritt_var,
            self.mock_status_label,
            self.mock_ergebnisse_text,
            self.mock_ergebnisse_button
        )

        # Check that the "image1.png" file was NOT overwritten
        with open(self.image1_path, "r") as f:
            content_after_run = f.read()

        self.assertEqual(self.image1_original_content, content_after_run, "The file 'image1.png' was overwritten.")

        # Check that an error message was shown
        mock_messagebox.showerror.assert_called_once()

if __name__ == '__main__':
    unittest.main()