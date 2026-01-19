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
    RevitLinkInstance,
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
        assign_workset(doc, new_ids, active_workset_id)
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
    if active_workset_name:
        lines.append("Target workset: {}".format(active_workset_name))
    forms.alert("\n".join(lines))


if __name__ == "__main__":
    main()
