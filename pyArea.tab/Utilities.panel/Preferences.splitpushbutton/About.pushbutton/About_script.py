# -*- coding: utf-8 -*-
"""Show the installed pyArea version, host pyRevit version and project links."""

__title__ = "About"
__author__ = "pyArea"

import os
import sys
import json

import clr

from pyrevit import script

# Add lib to path (this button sits one level deeper than most, inside the
# Preferences splitpushbutton, so lib is three levels up)
script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from _version import __version__, __release_date__

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System.Windows
from System.Windows import Window, Thickness, WindowStartupLocation, CornerRadius
from System.Windows.Controls import StackPanel, TextBlock, Border, WrapPanel, Orientation
from System.Windows.Input import Cursors
from System.Windows.Media import Brushes, SolidColorBrush, Color, FontFamily
from System.Windows.Media.Effects import DropShadowEffect

# Segoe MDL2 Assets ships with Windows 10/11 - "copy" and "checkmark" glyphs
ICON_FONT = "Segoe MDL2 Assets"
COPY_GLYPH = u""
CHECK_GLYPH = u""

# extension root is four levels up from this script
EXT_ROOT = os.path.abspath(os.path.join(script_dir, "..", "..", "..", ".."))

ACCENT = SolidColorBrush(Color.FromRgb(62, 74, 87))
MUTED = SolidColorBrush(Color.FromRgb(120, 128, 138))
HOVER = SolidColorBrush(Color.FromRgb(90, 104, 120))


def get_manifest():
    """Read author and repo url from extension.json.

    Keeps this dialog from duplicating metadata that already lives in the
    manifest - same reason the version lives only in lib/_version.py.
    """
    defaults = {
        "name": "pyArea",
        "author": "",
        "url": "",
    }
    try:
        manifest_path = os.path.join(EXT_ROOT, "extension.json")
        with open(manifest_path, "r") as manifest_file:
            data = json.load(manifest_file)
        for key in defaults:
            if data.get(key):
                defaults[key] = data[key]
    except Exception:
        pass
    return defaults


def get_pyrevit_version():
    """Return the host pyRevit version as 'vX.Y.Z', or None if unavailable."""
    try:
        from pyrevit import versionmgr
        return 'v{}'.format(
            versionmgr.get_pyrevit_version().get_formatted(strict=True))
    except Exception:
        return None


def get_revit_version():
    """Return the full host Revit version, e.g. '2026.4.20'.

    HOST_APP.subversion wraps Revit's own SubVersionNumber, which carries the
    update level - HOST_APP.version would only give '2026'.
    """
    try:
        from pyrevit import HOST_APP
        return HOST_APP.subversion
    except Exception:
        return None


def build_debug_info():
    """Environment summary users can paste when reporting a problem.

    Deliberately richer than what the dialog shows - the host and engine
    details are noise on screen but are the first things worth knowing in a
    bug report.
    """
    lines = [
        "pyArea {} (released {})".format(__version__, __release_date__),
    ]

    pyrvt_ver = get_pyrevit_version()
    if pyrvt_ver:
        lines.append("pyRevit {}".format(pyrvt_ver))

    try:
        from pyrevit import HOST_APP
        lines.append("Revit {} (build {})".format(
            HOST_APP.version, HOST_APP.build))
    except Exception:
        pass

    try:
        from pyrevit.userconfig import user_config
        cpyver = user_config.get_active_cpython_engine()
        runtime = sys.version.split('(')[0].strip()
        if cpyver:
            lines.append("Engine {} (cpython {})".format(
                runtime, cpyver.Version))
        else:
            lines.append("Engine {}".format(runtime))
    except Exception:
        pass

    return "\n".join(lines)


def copy_to_clipboard(text):
    """Copy text, returning True on success.

    WPF's Clipboard can intermittently fail inside Revit, so fall back to the
    WinForms implementation before giving up.
    """
    try:
        from System.Windows import Clipboard
        Clipboard.SetDataObject(text, True)
        return True
    except Exception:
        pass

    try:
        clr.AddReference('System.Windows.Forms')
        from System.Windows.Forms import Clipboard as FormsClipboard
        FormsClipboard.SetText(text)
        return True
    except Exception:
        return False


class AboutWindow(Window):
    """Borderless-feeling About card for pyArea."""

    def __init__(self):
        self.manifest = get_manifest()

        self.Width = 460
        self.SizeToContent = System.Windows.SizeToContent.Height
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        # Chromeless + transparent so the rounded card and its shadow can be
        # drawn by the content Border, same approach as pyRevit's own About.
        # getattr because "None" is a Python keyword and cannot be an attribute.
        self.WindowStyle = getattr(System.Windows.WindowStyle, "None")
        self.AllowsTransparency = True
        self.Background = Brushes.Transparent

        # No title bar means no close button: click anywhere or press a key
        self.MouseLeftButtonUp += lambda s, a: self.Close()
        self.KeyDown += lambda s, a: self.Close()

        self._create_ui()

    def _create_ui(self):
        panel = StackPanel()
        panel.Margin = Thickness(20, 22, 20, 18)

        panel.Children.Add(self._title_row())
        panel.Children.Add(self._info_block())
        panel.Children.Add(self._links_row())
        panel.Children.Add(self._footer())

        shadow = DropShadowEffect()
        shadow.Color = Color.FromRgb(45, 52, 61)
        shadow.BlurRadius = 15
        shadow.ShadowDepth = 2
        shadow.Opacity = 0.25

        card = Border()
        card.Background = Brushes.White
        card.CornerRadius = CornerRadius(12)
        # Margin leaves room for the shadow to render outside the card
        card.Margin = Thickness(20)
        card.Effect = shadow
        card.Child = panel

        self.Content = card

    def _title_row(self):
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        row.Margin = Thickness(0, 0, 0, 12)

        name = TextBlock()
        name.Text = self.manifest["name"]
        name.FontSize = 34
        name.Foreground = ACCENT
        name.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
        row.Children.Add(name)

        ver = TextBlock()
        ver.Text = " v{}".format(__version__)
        ver.FontSize = 14
        ver.Foreground = MUTED
        ver.Margin = Thickness(6, 0, 0, 6)
        ver.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
        row.Children.Add(ver)

        return row

    def _info_block(self):
        panel = StackPanel()
        panel.Margin = Thickness(0, 0, 0, 14)

        panel.Children.Add(self._release_row())

        # pyRevit and Revit share one line, sitting directly under the release date
        host_parts = []
        pyrvt_ver = get_pyrevit_version()
        if pyrvt_ver:
            host_parts.append("pyRevit {}".format(pyrvt_ver))

        revit_ver = get_revit_version()
        if revit_ver:
            host_parts.append("Revit {}".format(revit_ver))

        if host_parts:
            panel.Children.Add(self._info_line(
                "Running on {}".format("  ·  ".join(host_parts))))

        return panel

    def _info_line(self, text):
        line = TextBlock()
        line.Text = text
        line.FontSize = 11
        line.Foreground = MUTED
        line.TextAlignment = System.Windows.TextAlignment.Center
        line.Margin = Thickness(0, 1, 0, 1)
        return line

    def _release_row(self):
        """Release date with the copy-debug-info icon sitting after it."""
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = System.Windows.HorizontalAlignment.Center

        row.Children.Add(
            self._info_line("Released {}".format(__release_date__)))
        row.Children.Add(self._copy_icon())

        return row

    def _copy_icon(self):
        icon = TextBlock()
        icon.Text = COPY_GLYPH
        icon.FontFamily = FontFamily(ICON_FONT)
        icon.FontSize = 11
        icon.Foreground = MUTED
        icon.Margin = Thickness(6, 0, 0, 0)
        icon.VerticalAlignment = System.Windows.VerticalAlignment.Center
        icon.Cursor = Cursors.Hand
        icon.ToolTip = "Copy version info for bug reports"

        icon.MouseEnter += lambda s, a: setattr(s, 'Foreground', ACCENT)
        icon.MouseLeave += lambda s, a: setattr(s, 'Foreground', MUTED)
        icon.MouseLeftButtonUp += self._on_copy

        self._copy_glyph = icon
        return icon

    def _on_copy(self, sender, args):
        copied = copy_to_clipboard(build_debug_info())
        self._copy_glyph.Text = CHECK_GLYPH if copied else COPY_GLYPH
        self._copy_glyph.ToolTip = "Copied" if copied else "Copy failed"
        # Stop the click bubbling to the window so the dialog stays open and
        # the user actually sees the confirmation
        args.Handled = True

    def _links_row(self):
        row = WrapPanel()
        row.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        row.Margin = Thickness(0, 0, 0, 14)

        repo = self.manifest["url"].replace(".git", "")
        links = [("GitHub", repo)]
        if repo:
            links.append(("Releases", repo + "/releases"))
            links.append(("Report an Issue", repo + "/issues"))

        for label, url in links:
            if url:
                row.Children.Add(self._link_pill(label, url))

        return row

    def _make_pill(self, label):
        """A rounded chip in the style of pyRevit's About. Returns (chip, text)."""
        pill = Border()
        # Border needs a non-null Background to be hit-testable
        pill.Background = ACCENT
        pill.CornerRadius = CornerRadius(9)
        pill.Padding = Thickness(12, 4, 12, 5)
        pill.Margin = Thickness(3, 3, 3, 3)
        pill.Cursor = Cursors.Hand

        text = TextBlock()
        text.Text = label
        text.FontSize = 11
        text.Foreground = Brushes.White
        pill.Child = text

        pill.MouseEnter += lambda s, a: setattr(s, 'Background', HOVER)
        pill.MouseLeave += lambda s, a: setattr(s, 'Background', ACCENT)

        return pill, text

    def _link_pill(self, label, url):
        pill, _ = self._make_pill(label)
        # Click bubbles up to the window, which closes it - intended here
        pill.MouseLeftButtonUp += lambda s, a: script.open_url(url)
        return pill

    def _footer(self):
        footer = TextBlock()
        author = self.manifest["author"]
        footer.Text = u"© {}".format(author) if author else u"©"
        footer.FontSize = 10
        footer.Foreground = MUTED
        footer.TextAlignment = System.Windows.TextAlignment.Center
        return footer


AboutWindow().ShowDialog()
