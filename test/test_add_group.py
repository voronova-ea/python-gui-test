from model.group import Group


def test_add_group(app, excel_groups):
    old_list = app.groups.get_group_list()
    group = excel_groups
    app.groups.add_new_group(group)
    new_list = app.groups.get_group_list()
    old_list.append(group)
    assert sorted(old_list, key=Group.name) == sorted(new_list, key=Group.name)