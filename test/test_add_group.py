# def test_add_group(app, xlsx_groups):
def test_add_group(app):
    group = "my group"
    old_list = app.groups.get_group_list()
    app.groups.add_new_group(group)
    new_list = app.groups.get_group_list()
    assert len(old_list) + 1 == len(new_list)
    old_list.append(group)
    assert sorted(old_list) == sorted(new_list)
    # pass