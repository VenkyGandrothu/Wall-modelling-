import clr, sys, os, math, re
from Autodesk.Revit.DB import (
    ImportInstance, Options, GeometryElement, GeometryInstance,
    Line, Arc, PolyLine, Mesh, Solid, Family, FamilySymbol, BuiltInCategory,
    BuiltInParameter, FilteredElementCollector, Level, XYZ, Transaction, Wall,
    WallType, WallKind, WallUtils, JoinGeometryUtils, ElementTransformUtils)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import script
import Autodesk.Revit.DB as DB

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
import System.Windows.Forms as WinForms
from System.Windows import Window, Thickness
from System.Windows.Markup import XamlReader
import System.Windows.Controls

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# --- Setup Paths, Icon and Colors ---
script_path = script.get_script_path()
script_dir = os.path.dirname(script_path)
parent_dir = os.path.dirname(script_dir)
icon_path = os.path.join(parent_dir, "Resources", "10D.ico")
icon_path_escaped = icon_path.replace("\\", "\\\\")
COLOR_BG = "#FFFFFF"
COLOR_BORDER = "#35ADB8"
COLOR_ACCENT = "#35ADB8"
COLOR_TEXT = "#000000"

# --- CAD File Selection ---
class CADFileSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, ImportInstance)
    def AllowReference(self, ref, point):
        return False

try:
    selected_ref = uidoc.Selection.PickObject(ObjectType.Element, CADFileSelectionFilter(), "Select a CAD file")
    selected_cad_link = doc.GetElement(selected_ref)
except Exception:
    sys.exit()

cadFileName = selected_cad_link.Name

def get_layer_names(cad_link):
    names = set()
    cat = cad_link.Category
    if cat and cat.SubCategories:
        for sub in cat.SubCategories:
            names.add(sub.Name)
    return sorted(names)

all_layers = get_layer_names(selected_cad_link)
if not all_layers:
    sys.exit()

def get_elements_coordinates_by_layer(doc, selected_layer, cad_link):
    if not doc or not cad_link:
        return ([], [])
    opt = Options()
    opt.IncludeNonVisibleObjects = True
    opt.ComputeReferences = True
    geom_elem = cad_link.get_Geometry(opt)
    if geom_elem is None:
        return ([], [])
    elements = []
    blocks = []
    def extract(geom):
        try:
            for sub_geom in geom:
                extract(sub_geom)
            return
        except TypeError:
            pass
        if isinstance(geom, GeometryInstance):
            blocks.append(geom)
            instance_geom = geom.GetInstanceGeometry()
            if instance_geom:
                extract(instance_geom)
            return
        gs_id = getattr(geom, "GraphicsStyleId", None)
        if gs_id:
            gs = doc.GetElement(gs_id)
            if gs and gs.GraphicsStyleCategory and gs.GraphicsStyleCategory.Name.lower() == selected_layer.lower():
                geom_type = geom.GetType().Name.lower()
                if geom_type in ["mline", "multiline"]:
                    elements.append({
                        'start': geom.GetEndPoint(0) if hasattr(geom, "GetEndPoint") else None,
                        'end': geom.GetEndPoint(1) if hasattr(geom, "GetEndPoint") else None,
                        'type': 'MLine'
                    })
                elif geom_type == "line" or isinstance(geom, Line):
                    elements.append({
                        'start': geom.GetEndPoint(0),
                        'end': geom.GetEndPoint(1),
                        'type': 'Line'
                    })
                elif geom_type == "arc" or isinstance(geom, Arc):
                    elements.append({
                        'start': geom.GetEndPoint(0),
                        'end': geom.GetEndPoint(1),
                        'type': 'Arc'
                    })
                elif geom_type == "polyline" or isinstance(geom, PolyLine):
                    pts = geom.GetCoordinates()
                    elements.append({'points': pts, 'type': 'PolyLine'})
                elif isinstance(geom, Mesh):
                    for face in geom.Faces:
                        verts = face.GetVertices()
                        for i in range(len(verts)-1):
                            elements.append({
                                'start': verts[i],
                                'end': verts[i+1],
                                'type': 'Mesh Edge'
                            })
                elif isinstance(geom, Solid):
                    for face in geom.Faces:
                        verts = geom.Faces[0].GetVertices()
                        for i in range(len(verts)-1):
                            elements.append({
                                'start': verts[i],
                                'end': verts[i+1],
                                'type': 'Solid Edge'
                            })
        else:
            if hasattr(geom, "GetEndPoint"):
                try:
                    elements.append({
                        'start': geom.GetEndPoint(0),
                        'end': geom.GetEndPoint(1),
                        'type': geom.GetType().Name
                    })
                except Exception:
                    pass
    extract(geom_elem)
    return (elements, blocks)

all_wall_types = FilteredElementCollector(doc).OfClass(WallType).ToElements()
basic_wall_types = [wt for wt in all_wall_types if wt.Kind == WallKind.Basic]

def get_wall_type_display_name(wt):
    param = wt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    if param:
        type_name = param.AsString()
        if type_name and type_name.strip() != "":
            return "{0} - {1}".format(wt.FamilyName, type_name)
    return "{0} - {1}".format(getattr(wt, "FamilyName", "System Wall"), getattr(wt, "Name", "Unnamed"))

wall_names = sorted(set(get_wall_type_display_name(wt) for wt in basic_wall_types))
levels = sorted(FilteredElementCollector(doc).OfClass(Level).ToElements(), key=lambda lvl: lvl.Elevation)
level_names = [lvl.Name for lvl in levels]
try:
    active_level_name = doc.ActiveView.GenLevel.Name
except Exception:
    active_level_name = level_names[0] if level_names else ""
active_index = None
for i, lvl in enumerate(level_names):
    if lvl == active_level_name:
        active_index = i
        break
if active_index is None:
    active_index = 0
bottom_default = levels[active_index-1] if active_index > 0 else levels[active_index]

def find_family_symbol(wall_fam):
    for wt in basic_wall_types:
        if get_wall_type_display_name(wt) == wall_fam:
            return wt
    return None

def find_level_by_name(name):
    for lvl in levels:
        if lvl.Name == name:
            return lvl
    return None

def segment_direction(p1, p2):
    dx = p2.X - p1.X
    dy = p2.Y - p1.Y
    mag = math.sqrt(dx*dx + dy*dy)
    if mag < 1e-9:
        return (0, 0)
    return (dx/mag, dy/mag)

# --- Snapping Helper ---
def snap_angle(angle, tol=0.01):
    cardinals = [0, math.pi/2, math.pi, 3*math.pi/2]
    mod_angle = angle % (2*math.pi)
    for c in cardinals:
        if abs(mod_angle - c) < tol:
            return c
    return angle

# --- New Helper: Check if a wall already exists in the document at a given location ---
def wall_exists(candidate_line, tol=0.005):
    existing_walls = FilteredElementCollector(doc).OfClass(Wall).ToElements()
    for wall in existing_walls:
        try:
            locCurve = wall.Location.Curve
            if ((candidate_line.GetEndPoint(0).DistanceTo(locCurve.GetEndPoint(0)) < tol and
                 candidate_line.GetEndPoint(1).DistanceTo(locCurve.GetEndPoint(1)) < tol) or
                (candidate_line.GetEndPoint(0).DistanceTo(locCurve.GetEndPoint(1)) < tol and
                 candidate_line.GetEndPoint(1).DistanceTo(locCurve.GetEndPoint(0)) < tol)):
                return True
        except:
            continue
    return False

# --- UI Helpers ---
def attach_dropdown_open(combo):
    def open_dropdown(sender, e):
        sender.IsDropDownOpen = True
    combo.PreviewMouseLeftButtonDown += open_dropdown
    combo.GotFocus += open_dropdown

def filter_wall_family_items(combo):
    edt = combo.Template.FindName("PART_EditableTextBox", combo)
    if edt is None:
        return
    txt = edt.Text.strip().lower()
    combo.Items.Clear()
    for fam in wall_names:
        if fam.lower().startswith(txt):
            combo.Items.Add(fam)
    combo.IsDropDownOpen = True

# --- Build UI ---
selected_cad_link = doc.GetElement(selected_ref)
if selected_cad_link.Category:
    cad_file_name = selected_cad_link.Category.Name
else:
    cad_file_name = "Unknown CAD File"

# Updated XAML:
# - The Wall Families panel is wrapped in a ScrollViewer with MaxHeight to show up to four rows.
# - The ScrollViewer is given the name "FamilyScrollViewer".
xaml_str = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="10dimensionsdesign" SizeToContent="Height" Width="700"
        WindowStartupLocation="CenterScreen" Icon="{0}" Background="{1}">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0" Margin="0 10">
      <TextBlock Text="Selected CAD File:" Foreground="{2}" FontWeight="Bold" Margin="0,0,0,5"/>
      <TextBlock x:Name="CADFileNameTextBlock" Foreground="{2}" FontWeight="Normal" Margin="0,0,0,10"/>
    </StackPanel>
    <StackPanel Grid.Row="1" Margin="0 10">
      <TextBlock Text="Select a CAD Layer:" Foreground="{2}" FontWeight="Bold" Margin="0,0,0,10"/>
      <ComboBox x:Name="LayerComboBox" IsEditable="True" IsTextSearchEnabled="False" Height="25"
                Margin="0,0,0,10" BorderBrush="{3}" Foreground="Black" Background="{1}"/>
    </StackPanel>
    <StackPanel Grid.Row="2" Margin="0 10">
      <TextBlock Text="Wall Families:" Foreground="{2}" FontWeight="Bold" Margin="0,0,0,10"/>
      <ScrollViewer x:Name="FamilyScrollViewer" MaxHeight="120" VerticalScrollBarVisibility="Auto">
        <StackPanel x:Name="FamilyDropdownsPanel" Orientation="Vertical">
          <ComboBox x:Name="WallFamilyComboBox" IsEditable="True" IsTextSearchEnabled="True" Height="25" HorizontalAlignment="Stretch"
                    BorderBrush="{3}" Foreground="Black" Background="{1}"/>
        </StackPanel>
      </ScrollViewer>
      <Button x:Name="AddNewFamilyButton" Content="+ Add new family" Width="110" HorizontalAlignment="Right"
              Margin="0,5,0,0" Background="{3}" Foreground="White" FontWeight="Bold" Height="25"/>
    </StackPanel>
    <Grid Grid.Row="3" Margin="0 10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <StackPanel Grid.Column="0" Orientation="Vertical" HorizontalAlignment="Stretch">
        <TextBlock Text="Base Constraint:" Foreground="{2}" FontWeight="Bold" Margin="5"/>
        <ComboBox x:Name="BottomLevelComboBox" IsEditable="True" IsTextSearchEnabled="True" Height="25" Margin="5"
                  BorderBrush="{3}" Foreground="Black" Background="{1}"/>
      </StackPanel>
      <StackPanel Grid.Column="1" Orientation="Vertical" HorizontalAlignment="Stretch">
        <TextBlock Text="Base Offset:" Foreground="{2}" FontWeight="Bold" Margin="5"/>
        <TextBox x:Name="BottomOffsetTextBox" Height="25" Margin="5"
                 BorderBrush="{3}" Foreground="Black" Background="{1}" Text="0"
                 HorizontalContentAlignment="Center" TextAlignment="Center"/>
      </StackPanel>
    </Grid>
    <Grid Grid.Row="4" Margin="0 10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <StackPanel Grid.Column="0" Orientation="Vertical" HorizontalAlignment="Stretch">
        <TextBlock Text="Top Constraint:" Foreground="{2}" FontWeight="Bold" Margin="5"/>
        <ComboBox x:Name="TopLevelComboBox" IsEditable="True" IsTextSearchEnabled="True" Height="25" Margin="5"
                  BorderBrush="{3}" Foreground="Black" Background="{1}"/>
      </StackPanel>
      <StackPanel Grid.Column="1" Orientation="Vertical" HorizontalAlignment="Stretch">
        <TextBlock Text="Top Offset:" Foreground="{2}" FontWeight="Bold" Margin="5"/>
        <TextBox x:Name="TopZOffsetTextBox" Height="25" Margin="5"
                 BorderBrush="{3}" Foreground="Black" Background="{1}" Text="0"
                 HorizontalContentAlignment="Center" TextAlignment="Center"/>
      </StackPanel>
    </Grid>
    <StackPanel Grid.Row="5" Margin="0 10" Orientation="Horizontal">
      <CheckBox x:Name="DisallowJoinCheckBox" Content="Disallow Wall Joins" Foreground="{2}"/>
    </StackPanel>
    <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Center" Margin="10">
      <Button x:Name="OKButton" Content="OK" Width="80" Margin="5">
        <Button.Style>
          <Style TargetType="Button">
            <Setter Property="Background" Value="{3}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Style.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="White"/>
                <Setter Property="Foreground" Value="{3}"/>
              </Trigger>
            </Style.Triggers>
          </Style>
        </Button.Style>
      </Button>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" Margin="5">
        <Button.Style>
          <Style TargetType="Button">
            <Setter Property="Background" Value="{3}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Style.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="White"/>
                <Setter Property="Foreground" Value="{3}"/>
              </Trigger>
            </Style.Triggers>
          </Style>
        </Button.Style>
      </Button>
    </StackPanel>
  </Grid>
</Window>
""".format(icon_path_escaped, COLOR_BG, COLOR_TEXT, COLOR_BORDER)

window = XamlReader.Parse(xaml_str)
cad_file_text_block = window.FindName("CADFileNameTextBlock")
if cad_file_text_block is not None:
    cad_file_text_block.Text = cad_file_name

layer_combo = window.FindName("LayerComboBox")
for layer in all_layers:
    layer_combo.Items.Add(layer)
attach_dropdown_open(layer_combo)

wall_family_combo = window.FindName("WallFamilyComboBox")
wall_family_combo.Items.Clear()
for fam in wall_names:
    wall_family_combo.Items.Add(fam)
attach_dropdown_open(wall_family_combo)
wall_family_combo.ApplyTemplate()
edt_wall = wall_family_combo.Template.FindName("PART_EditableTextBox", wall_family_combo)
if edt_wall is not None:
    def on_wall_family_text_changed(sender, e):
        filter_wall_family_items(wall_family_combo)
    edt_wall.TextChanged += on_wall_family_text_changed

bottom_level_combo = window.FindName("BottomLevelComboBox")
top_level_combo = window.FindName("TopLevelComboBox")
for lvl in level_names:
    bottom_level_combo.Items.Add(lvl)
    top_level_combo.Items.Add(lvl)
for i, lvl in enumerate(level_names):
    if lvl == active_level_name:
        top_level_combo.SelectedIndex = i
        break
if active_index is not None and active_index > 0:
    bottom_level_combo.SelectedItem = levels[active_index - 1].Name
else:
    bottom_level_combo.SelectedItem = active_level_name
attach_dropdown_open(bottom_level_combo)
attach_dropdown_open(top_level_combo)

def filter_combo_items():
    edt = layer_combo.Template.FindName("PART_EditableTextBox", layer_combo)
    if edt is None:
        return
    txt = edt.Text
    layer_combo.Items.Clear()
    for layer in all_layers:
        if layer.lower().startswith(txt.lower()):
            layer_combo.Items.Add(layer)
    layer_combo.IsDropDownOpen = True

def on_layer_text_changed(sender, e):
    filter_combo_items()

layer_combo.ApplyTemplate()
edt = layer_combo.Template.FindName("PART_EditableTextBox", layer_combo)
if edt is not None:
    edt.TextChanged += on_layer_text_changed

family_dropdowns_panel = window.FindName("FamilyDropdownsPanel")
add_new_family_btn = window.FindName("AddNewFamilyButton")

def on_add_new_family_click(sender, e):
    new_combo = System.Windows.Controls.ComboBox()
    new_combo.Height = wall_family_combo.Height
    new_combo.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
    new_combo.IsEditable = True
    new_combo.IsTextSearchEnabled = True
    new_combo.Foreground = wall_family_combo.Foreground
    new_combo.Background = wall_family_combo.Background
    new_combo.Margin = Thickness(0, 5, 0, 0)
    new_combo.Items.Clear()
    for fam in wall_names:
        new_combo.Items.Add(fam)
    attach_dropdown_open(new_combo)
    new_combo.ApplyTemplate()
    if new_combo.Template is not None:
        edt_new = new_combo.Template.FindName("PART_EditableTextBox", new_combo)
        if edt_new is not None:
            def on_new_wall_family_text_changed(sender, e):
                filter_wall_family_items(new_combo)
            edt_new.TextChanged += on_new_wall_family_text_changed
    family_dropdowns_panel.Children.Add(new_combo)
    # After adding, scroll the FamilyScrollViewer to the end
    scrollViewer = window.FindName("FamilyScrollViewer")
    if scrollViewer is not None:
        scrollViewer.ScrollToEnd()

add_new_family_btn.Click += on_add_new_family_click

progress_xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Progress" Height="150" Width="400"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Icon="{0}" Background="White">
  <StackPanel Margin="10">
    <ProgressBar x:Name="pb" Minimum="0" Maximum="100" Value="0" Height="30" Foreground="{1}" Background="White"/>
    <TextBlock x:Name="txtProgress" Margin="0,10,0,0" HorizontalAlignment="Center" Foreground="{1}" FontWeight="Bold" FontSize="14"/>
    <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="25" HorizontalAlignment="Right" Margin="0,10,0,0" Background="{1}" Foreground="White"/>
  </StackPanel>
</Window>
""".format(icon_path_escaped, COLOR_ACCENT)

def on_ok_button_click(sender, e):
    sel_layer = layer_combo.SelectedItem if layer_combo.SelectedItem else layer_combo.Text
    if not sel_layer:
        return
    selected_families = []
    for child in family_dropdowns_panel.Children:
        fam = child.SelectedItem if child.SelectedItem else child.Text
        if fam:
            selected_families.append(fam)
    bottom_level_name = bottom_level_combo.SelectedItem if bottom_level_combo.SelectedItem else bottom_level_combo.Text
    try:
        base_offset = int(window.FindName("BottomOffsetTextBox").Text)
    except Exception:
        base_offset = 0
    top_level_name = top_level_combo.SelectedItem if top_level_combo.SelectedItem else top_level_combo.Text
    try:
        top_offset = int(window.FindName("TopZOffsetTextBox").Text)
    except Exception:
        top_offset = 0
    disallow_joins = window.FindName("DisallowJoinCheckBox").IsChecked

    lines_data, blocks_data = get_elements_coordinates_by_layer(doc, str(sel_layer), selected_cad_link)
    placements = []

    for element in lines_data:
        if element["type"].lower() not in ["polyline", "mline", "multiline"]:
            continue
        pts = element.get("points")
        if pts is None:
            if element.get("start") and element.get("end"):
                pts = [element["start"], element["end"]]
            else:
                continue
        if len(pts) < 2:
            continue

        segments = []
        for i in range(len(pts) - 1):
            p_start = pts[i]
            p_end = pts[i + 1]
            d = segment_direction(p_start, p_end)
            seg_length = math.sqrt((p_end.X - p_start.X) ** 2 + (p_end.Y - p_start.Y) ** 2)
            segments.append((p_start, p_end, d, seg_length))
        if len(segments) < 2:
            continue

        segments.sort(key=lambda s: s[3], reverse=True)
        seg1 = segments[0]
        seg2 = segments[1]

        d1 = seg1[2]
        d2 = seg2[2]
        if abs(d1[0] * d2[0] + d1[1] * d2[1]) < 0.99:
            continue
        diff_x = seg1[0].X - seg2[0].X
        diff_y = seg1[0].Y - seg2[0].Y
        perp_dot = abs(diff_x * (-d2[1]) + diff_y * d2[0])
        distance_ft = perp_dot
        distance_mm_int = int(round(distance_ft * 304.8, 0))

        proj1 = seg1[0].X * d1[0] + seg1[0].Y * d1[1]
        proj2 = seg1[1].X * d1[0] + seg1[1].Y * d1[1]
        seg1_min = min(proj1, proj2)
        seg1_max = max(proj1, proj2)
        projB1 = seg2[0].X * d1[0] + seg2[0].Y * d1[1]
        projB2 = seg2[1].X * d1[0] + seg2[1].Y * d1[1]
        seg2_min = min(projB1, projB2)
        seg2_max = max(projB1, projB2)
        overlap_start = max(seg1_min, seg2_min)
        overlap_end = min(seg1_max, seg2_max)
        overlap_length = overlap_end - overlap_start
        if overlap_length <= 0:
            continue

        candidate_data = None

        for wall_fam in selected_families:
            extracted_ints = [int(x) for x in re.findall(r'\d+', wall_fam) if int(x) > 10]
            if not extracted_ints:
                continue
            desired_width = extracted_ints[0]
            if (desired_width - 0.5) <= distance_mm_int <= (desired_width + 0.5):
                mid_proj = (overlap_start + overlap_end) / 2.0
                dot_p1 = seg1[0].X * d1[0] + seg1[0].Y * d1[1]
                offset1 = mid_proj - dot_p1
                point_on_line1 = XYZ(seg1[0].X + d1[0] * offset1,
                                     seg1[0].Y + d1[1] * offset1,
                                     seg1[0].Z)
                dot_p2 = seg2[0].X * d1[0] + seg2[0].Y * d1[1]
                offset2 = mid_proj - dot_p2
                point_on_line2 = XYZ(seg2[0].X + d1[0] * offset2,
                                     seg2[0].Y + d1[1] * offset2,
                                     seg2[0].Z)
                mid_point = XYZ((point_on_line1.X + point_on_line2.X) / 2.0,
                                (point_on_line1.Y + point_on_line2.Y) / 2.0,
                                (point_on_line1.Z + point_on_line2.Z) / 2.0)
                angle = math.atan2(d1[1], d1[0])
                angle = snap_angle(angle, tol=0.01)
                candidate_data = (mid_point, angle, overlap_length, wall_fam)
                break
        if candidate_data:
            placements.append(candidate_data)

    deduped_placements = []
    for candidate in placements:
        duplicate = False
        for existing in deduped_placements:
            if (candidate[0].DistanceTo(existing[0]) < 0.01 and
                abs(candidate[1]-existing[1]) < 0.01 and
                candidate[3] == existing[3]):
                duplicate = True
                break
        if not duplicate:
            deduped_placements.append(candidate)
    placements = deduped_placements

    bottom_elev = None
    top_elev = None
    for lvl in levels:
        if lvl.Name == bottom_level_name:
            bottom_elev = lvl.Elevation
        if lvl.Name == top_level_name:
            top_elev = lvl.Elevation
    if bottom_elev is None:
        bottom_elev = 0
    if top_elev is None:
        top_elev = 0
    level_distance = abs(top_elev - bottom_elev)
    total_wall_count = len(placements)

    window.Close()

    progress_window = XamlReader.Parse(progress_xaml)
    pb = progress_window.FindName("pb")
    txtProgress = progress_window.FindName("txtProgress")
    btnCancel = progress_window.FindName("btnCancel")
    pb.Maximum = total_wall_count
    cancelled = [False]
    def on_cancel_progress(sender, e):
        cancelled[0] = True
        progress_window.Close()
    btnCancel.Click += on_cancel_progress
    progress_window.Show()

    current = 0
    created_walls = []  # store (wall, wall_line)
    try:
        t_place = Transaction(doc, "Place Walls")
        t_place.Start()
    except Exception:
        return
    tolerance = doc.Application.ShortCurveTolerance
    unique_fams = set([wf for (_, _, _, wf) in placements])
    for wf in unique_fams:
        fs_temp = find_family_symbol(wf)
        if fs_temp and hasattr(fs_temp, "IsActive") and not fs_temp.IsActive:
            try:
                t_activate = Transaction(doc, "Activate Wall Family")
                t_activate.Start()
                fs_temp.Activate()
                doc.Regenerate()
                t_activate.Commit()
            except Exception:
                pass
    topLevel = doc.GetElement(find_level_by_name(top_level_name).Id) if find_level_by_name(top_level_name) else None
    baseLevel = doc.GetElement(find_level_by_name(bottom_level_name).Id) if find_level_by_name(bottom_level_name) else None
    for (pt, angle, wall_length, wall_fam) in placements:
        if cancelled[0]:
            break
        fs = find_family_symbol(wall_fam)
        if fs is None:
            continue
        fs = doc.GetElement(fs.Id)
        if wall_length < tolerance:
            continue
        direction_vector = XYZ(math.cos(angle), math.sin(angle), 0)
        start_pt = XYZ(pt.X - (wall_length/2)*direction_vector.X,
                       pt.Y - (wall_length/2)*direction_vector.Y,
                       pt.Z)
        end_pt = XYZ(pt.X + (wall_length/2)*direction_vector.X,
                     pt.Y + (wall_length/2)*direction_vector.Y,
                     pt.Z)
        try:
            wall_line = DB.Line.CreateBound(start_pt, end_pt)
            actual_dir = XYZ(end_pt.X - start_pt.X, end_pt.Y - start_pt.Y, end_pt.Z - start_pt.Z)
            dot_dir = actual_dir.X * direction_vector.X + actual_dir.Y * direction_vector.Y
            if dot_dir < 0:
                start_pt, end_pt = end_pt, start_pt
                wall_line = DB.Line.CreateBound(start_pt, end_pt)
        except Exception:
            continue
        if wall_line.Length < tolerance:
            continue
        duplicate = False
        for (existing_wall, existing_line) in created_walls:
            if (wall_line.GetEndPoint(0).DistanceTo(existing_line.GetEndPoint(0)) < 0.005 and
                wall_line.GetEndPoint(1).DistanceTo(existing_line.GetEndPoint(1)) < 0.005):
                duplicate = True
                break
        if not duplicate and wall_exists(wall_line, tol=0.005):
            duplicate = True
        if duplicate:
            continue
        level = find_level_by_name(bottom_level_name)
        if level is None:
            level = levels[0]
        level = doc.GetElement(level.Id)
        base_offset_ft = base_offset/304.8
        wall_height = level_distance + (top_offset/304.8) - base_offset_ft
        try:
            new_wall = Wall.Create(doc, wall_line, fs.Id, level.Id, wall_height, base_offset_ft, False, False)
        except Exception:
            continue
        allow_disallow = True
        for (existing_wall, existing_line) in created_walls:
            if new_wall.WallType.Id == existing_wall.WallType.Id:
                if (wall_line.GetEndPoint(0).DistanceTo(existing_line.GetEndPoint(0)) < 0.005 and
                    wall_line.GetEndPoint(1).DistanceTo(existing_line.GetEndPoint(1)) < 0.005):
                    allow_disallow = False
                    break
        if disallow_joins and allow_disallow:
            try:
                WallUtils.DisallowWallJoinAtEnd(new_wall, 0)
                WallUtils.DisallowWallJoinAtEnd(new_wall, 1)
            except Exception:
                pass

        created_walls.append((new_wall, wall_line))
        try:
            if baseLevel is not None:
                param_base = new_wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
                if param_base:
                    param_base.Set(baseLevel.Id)
            if topLevel is not None:
                param_top = new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
                if param_top:
                    param_top.Set(topLevel.Id)
        except Exception:
            pass
        try:
            param_base_offset = new_wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
            if param_base_offset:
                param_base_offset.Set(base_offset_ft)
            param_top_offset = new_wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
            if param_top_offset:
                param_top_offset.Set(top_offset/304.8)
        except Exception:
            pass
        current += 1
        txtProgress.Text = "Walls Created: {0} / {1}".format(current, total_wall_count)
        pb.Value = float(current)
        WinForms.Application.DoEvents()
    try:
        t_place.Commit()
    except Exception:
        try:
            t_place.RollBack()
        except Exception:
            pass
    try:
        t_final = Transaction(doc, "Finalize")
        t_final.Start()
        t_final.Commit()
    except Exception:
        pass
    progress_window.Close()

def wall_exists(candidate_line, tol=0.005):
    existing_walls = FilteredElementCollector(doc).OfClass(Wall).ToElements()
    for wall in existing_walls:
        try:
            locCurve = wall.Location.Curve
            if ((candidate_line.GetEndPoint(0).DistanceTo(locCurve.GetEndPoint(0)) < tol and
                 candidate_line.GetEndPoint(1).DistanceTo(locCurve.GetEndPoint(1)) < tol) or
                (candidate_line.GetEndPoint(0).DistanceTo(locCurve.GetEndPoint(1)) < tol and
                 candidate_line.GetEndPoint(1).DistanceTo(locCurve.GetEndPoint(0)) < tol)):
                return True
        except Exception:
            continue
    return False

def find_level_by_name(name):
    for lvl in levels:
        if lvl.Name == name:
            return lvl
    return None

on_cancel_button_click = lambda sender, e: window.Close()
ok_button = window.FindName("OKButton")
ok_button.Click += on_ok_button_click
cancel_button = window.FindName("CancelButton")
if cancel_button is not None:
    cancel_button.Click += on_cancel_button_click

window.ShowDialog()
