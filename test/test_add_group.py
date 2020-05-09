def test_add_group(app):
    old_list = app.groups.get_group_list()
    app.groups.add_new_group("test group")
    new_list = app.groups.get_group_list()
    old_list.append("test")
    assert sorted(old_list) == sorted(new_list)