#! python
# pyRevit: Copy elements from a selected link by workset and category.

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CategoryType,
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementCategoryFilter,
    ElementId,
    ElementTransformUtils,
    ElementWorksetFilter,
    FilteredElementCollector,
    FilteredWorksetCollector,
    IDuplicateTypeNamesHandler,
    Level,
    RevitLinkInstance,
    StorageType,
    Transaction,
    WorksetKind,
)
from System.Collections.Generic import List

try:
    from pyrevit import forms
    uidoc = __revit__.ActiveUIDocument  # type: ignore[name-defined]
    doc = uidoc.Document
except Exception:
    raise Exception("Run from pyRevit inside Revit.")


class LinkWrap(object):
    def __init__(self, instance):
        self.instance = instance
        self.doc = instance.GetLinkDocument()
        self.title = self.doc.Title if self.doc else "Unloaded Link"
        self.name = instance.Name

    def __str__(self):
        if self.doc:
            return u"{} ({})".format(self.title, self.name)
        return u"{} (unloaded)".format(self.name)


class WorksetWrap(object):
    def __init__(self, workset):
        self.workset = workset
        self.name = workset.Name

    def __str__(self):
        return self.name


class CategoryWrap(object):
    def __init__(self, category, count):
        self.category = category
        self.count = count
        self.name = category.Name

    def __str__(self):
        return u"{} ({})".format(self.name, self.count)


class DuplicateTypeHandler(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


def _normalize_level_name(level_name):
    try:
        return level_name.strip().lower()
    except Exception:
        return None


def _get_level_param_bips():
    bips = [
        BuiltInParameter.LEVEL_PARAM,
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,
    ]
    for bip_name in [
        "INSTANCE_REFERENCE_LEVEL_PARAM",
        "FAMILY_BASE_LEVEL_PARAM",
        "FAMILY_TOP_LEVEL_PARAM",
        "RBS_START_LEVEL_PARAM",
        "RBS_END_LEVEL_PARAM",
        "RBS_REFERENCE_LEVEL_PARAM",
        "WALL_BASE_CONSTRAINT",
        "WALL_HEIGHT_TYPE",
    ]:
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip and bip not in bips:
            bips.append(bip)
    return bips


LEVEL_PARAM_BIPS = _get_level_param_bips()


def _build_level_id_map_by_name(host_doc):
    level_id_by_name = {}
    for level in FilteredElementCollector(host_doc).OfClass(Level):
        key = _normalize_level_name(level.Name)
        if key and key not in level_id_by_name:
            level_id_by_name[key] = level.Id
    return level_id_by_name


def _get_level_names_by_param(elem, elem_doc):
    names_by_bip = {}
    for bip in LEVEL_PARAM_BIPS:
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if not param or not param.HasValue:
            continue
        if param.StorageType != StorageType.ElementId:
            continue
        level_id = param.AsElementId()
        if level_id == ElementId.InvalidElementId:
            continue
        level_elem = elem_doc.GetElement(level_id)
        if not level_elem or not isinstance(level_elem, Level):
            continue
        names_by_bip[bip] = level_elem.Name
    return names_by_bip


def _remap_level_params_on_copied_elements(source_doc, dest_doc, source_ids, dest_ids):
    level_id_by_name = _build_level_id_map_by_name(dest_doc)
    if not level_id_by_name:
        return 0

    updated_element_count = 0
    for src_id, dst_id in zip(source_ids, dest_ids):
        src_elem = source_doc.GetElement(src_id)
        dst_elem = dest_doc.GetElement(dst_id)
        if not src_elem or not dst_elem:
            continue

        src_levels_by_bip = _get_level_names_by_param(src_elem, source_doc)
        if not src_levels_by_bip:
            continue

        updated_this_element = False
        for bip, level_name in src_levels_by_bip.items():
            dest_level_id = level_id_by_name.get(_normalize_level_name(level_name))
            if not dest_level_id:
                continue
            dst_param = dst_elem.get_Parameter(bip)
            if not dst_param or dst_param.IsReadOnly:
                continue
            if dst_param.StorageType != StorageType.ElementId:
                continue
            try:
                dst_param.Set(dest_level_id)
                updated_this_element = True
            except Exception:
                continue

        if updated_this_element:
            updated_element_count += 1

    return updated_element_count


def get_link_instances(host_doc):
    links = []
    for inst in FilteredElementCollector(host_doc).OfClass(RevitLinkInstance):
        if inst.GetLinkDocument():
            links.append(inst)
    return links


def get_user_worksets(link_doc):
    return list(FilteredWorksetCollector(link_doc).OfKind(WorksetKind.UserWorkset))


def collect_categories(link_doc, workset_id=None):
    collector = FilteredElementCollector(link_doc).WhereElementIsNotElementType()
    if workset_id:
        collector = collector.WherePasses(ElementWorksetFilter(workset_id))
    counts = {}
    for elem in collector:
        cat = elem.Category
        if not cat:
            continue
        if cat.CategoryType != CategoryType.Model:
            continue
        cid = cat.Id.IntegerValue
        if cid not in counts:
            counts[cid] = [cat, 0]
        counts[cid][1] += 1
    items = []
    for data in counts.values():
        items.append(CategoryWrap(data[0], data[1]))
    items.sort(key=lambda item: item.name.lower())
    return items


def collect_element_ids(link_doc, category, workset_id=None):
    collector = FilteredElementCollector(link_doc).WhereElementIsNotElementType()
    if workset_id:
        collector = collector.WherePasses(ElementWorksetFilter(workset_id))
    collector = collector.WherePasses(ElementCategoryFilter(category.Id))
    return [elem.Id for elem in collector]


def get_active_workset_id(host_doc):
    if not host_doc.IsWorkshared:
        return None
    return host_doc.GetWorksetTable().GetActiveWorksetId()


def assign_workset(host_doc, element_ids, workset_id):
    if not workset_id:
        return
    for eid in element_ids:
        elem = host_doc.GetElement(eid)
        if not elem:
            continue
        param = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        if param and not param.IsReadOnly:
            param.Set(workset_id.IntegerValue)


def main():
    links = get_link_instances(doc)
    if not links:
        forms.alert("No loaded Revit links found.", exitscript=True)

    link_pick = forms.SelectFromList.show(
        [LinkWrap(link) for link in links],
        title="Select link",
        multiselect=False
    )
    if not link_pick:
        forms.alert("No link selected.", exitscript=True)

    link_instance = link_pick.instance
    link_doc = link_pick.doc
    workset_id = None
    workset_name = None

    if link_doc.IsWorkshared:
        worksets = get_user_worksets(link_doc)
        if not worksets:
            forms.alert("No user worksets found in the linked model.", exitscript=True)
        workset_pick = forms.SelectFromList.show(
            [WorksetWrap(ws) for ws in worksets],
            title="Select workset",
            multiselect=False
        )
        if not workset_pick:
            forms.alert("No workset selected.", exitscript=True)
        workset_id = workset_pick.workset.Id
        workset_name = workset_pick.workset.Name
    else:
        forms.alert("Linked model is not workshared; workset filtering will be skipped.")

    categories = collect_categories(link_doc, workset_id)
    if not categories:
        forms.alert("No model categories found for the chosen filters.", exitscript=True)

    category_pick = forms.SelectFromList.show(
        categories,
        title="Select category",
        multiselect=False
    )
    if not category_pick:
        forms.alert("No category selected.", exitscript=True)

    category = category_pick.category
    element_ids = collect_element_ids(link_doc, category, workset_id)
    if not element_ids:
        forms.alert("No elements found for the selected link, workset, and category.", exitscript=True)

    element_id_list = List[ElementId]()
    for eid in element_ids:
        element_id_list.Add(eid)

    copy_options = CopyPasteOptions()
    copy_options.SetDuplicateTypeNamesHandler(DuplicateTypeHandler())

    transform = link_instance.GetTotalTransform()
    active_workset_id = get_active_workset_id(doc)
    active_workset_name = None
    if active_workset_id:
        try:
            active_workset_name = doc.GetWorksetTable().GetWorkset(active_workset_id).Name
        except Exception:
            active_workset_name = None

    transaction = Transaction(doc, "Copy Elements from Link")
    transaction.Start()
    try:
        new_ids = ElementTransformUtils.CopyElements(
            link_doc,
            element_id_list,
            doc,
            transform,
            copy_options
        )
        new_ids_list = list(new_ids)
        assign_workset(doc, new_ids_list, active_workset_id)
        level_remap_count = _remap_level_params_on_copied_elements(
            link_doc, doc, element_ids, new_ids_list
        )
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        forms.alert("Copy failed:\n{}".format(exc))
        raise

    lines = []
    lines.append("Link: {}".format(link_doc.Title))
    if workset_name:
        lines.append("Workset: {}".format(workset_name))
    lines.append("Category: {}".format(category.Name))
    lines.append("Copied: {}".format(len(new_ids)))
    lines.append("Levels remapped: {}".format(level_remap_count))
    if active_workset_name:
        lines.append("Target workset: {}".format(active_workset_name))
    forms.alert("\n".join(lines))


if __name__ == "__main__":
    main()
