from outlook_exporter import outlook


def test_search_mail_filters():
    results = outlook.search_mail("keyword")
    assert all("keyword" in s.lower() for s in results)
