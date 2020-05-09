from model.group import Group
import random


def test_del_group(app):
    app.groups.check_group_count(Group(name="some group"))
    old_list = app.groups.get_group_list()
    index = random.choice(range(len(old_list)))
    app.groups.del_group_by_index(index)
    new_list = app.groups.get_group_list()
    old_list.pop(index)
    assert sorted(old_list, key=Group.name) == sorted(new_list, key=Group.name)