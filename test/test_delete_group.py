def test_delete_group_by_name(app):
    if len(app.groups.get_group_list()) <= 1:
        app.groups.add_new_group("my group new")
    group = "my group new"
    old_list = app.groups.get_group_list()
    if group not in old_list:
        app.groups.add_new_group(group)
    old_list = app.groups.get_group_list()
    app.groups.delete_group(group)
    new_list = app.groups.get_group_list()
    assert len(old_list) - 1 == len(new_list)
    old_list.remove(group)
    assert sorted(old_list) == sorted(new_list)