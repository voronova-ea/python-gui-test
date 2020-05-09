from model.group import Group


class GroupHelper:

    def __init__(self, app):
        self.app = app

    def get_group_list(self):
        group_list = []
        self.open_group_editor()
        tree = self.group_editor.window(auto_id="uxAddressTreeView")
        root = tree.tree_root()
        for node in root.children():
            name = node.text()
            group_list.append(Group(name=name))
        self.close_group_editor()
        return group_list

    def add_new_group(self, group):
        self.open_group_editor()
        self.group_editor.window(auto_id="uxNewAddressButton").click()
        input = self.group_editor.window(class_name="Edit")
        input.set_text(group.name)
        input.type_keys("\n")
        self.close_group_editor()

    def open_group_editor(self):
        self.app.main_window.window(auto_id="groupButton").click()
        self.group_editor = self.app.application.window(title="Group editor")
        self.group_editor.wait("visible")

    def close_group_editor(self):
        self.group_editor.close()