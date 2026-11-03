"""HTTP security and per-site persistence checks. Requires Flask/Werkzeug.

These run in the complete Windows build environment; without the HTTP packages
we skip rather than reporting an unexecuted integration check as a pass.
"""
import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from redink_server.config import initialize, load_settings
from redink_server.settings_ui import update_configuration


@unittest.skipUnless(importlib.util.find_spec('flask'), 'Flask HTTP dependency unavailable')
class SharedSiteHttpTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.data = self.root/'data'
        initialize(self.data)
        self.site = self.root/'public'
        self.site.mkdir()
        (self.site/'index.html').write_text('<h1>Portal</h1>', encoding='utf-8')
        (self.site/'app.js').write_text('console.log("portal");', encoding='utf-8')
        self.settings = load_settings(self.data)

    def site_definition(self, name, path):
        return dict(name=name, enabled=True, host='127.0.0.1', port=8765, path=path,
                    directory=str(self.site), https=False, spa_fallback=False)

    def test_shared_wsgi_port_separates_static_routes_and_admin_auth(self):
        from werkzeug.test import Client
        from werkzeug.wrappers import Response
        from redink_server.web import create_app
        from redink_server.static_websites import create_all_listeners
        site=self.site_definition('Portal','/apps')
        update_configuration(self.settings,{'websites':{'sites':[site]}})
        admin=create_app(load_settings(self.data))
        admin.testing=True
        applications=[]
        def factory(settings, app):
            applications.append((settings.port,app))
            class Listener:
                pass
            return Listener()
        with patch('redink_server.static_websites.create_http_server',side_effect=factory):
            group=create_all_listeners(load_settings(self.data),admin)
        self.assertEqual(len(group._listeners),1)
        self.assertEqual(len(applications),1)
        client=Client(applications[0][1],Response)
        self.assertEqual(client.get('/').status_code,302)
        self.assertEqual(client.get('/login').status_code,200)
        self.assertEqual(client.get('/api/settings').status_code,401)
        self.assertEqual(client.post('/collect/invalid',json={}).status_code,401)
        self.assertEqual(client.get('/apps/').status_code,200)
        self.assertIn('Portal',client.get('/apps/').get_data(as_text=True))
        self.assertEqual(client.get('/apps/app.js').status_code,200)
        self.assertEqual(client.get('/apps/../server.toml').status_code,404)
        self.assertEqual(client.get('/apps-misroute').status_code,404)

    def test_per_site_api_saves_only_one_and_rejects_anonymous_or_bad_csrf(self):
        from redink_server.web import create_app
        app=create_app(self.settings)
        app.testing=True
        client=app.test_client()
        portal=self.site_definition('Portal','/apps')
        self.assertEqual(client.post('/api/settings/website',json={'action':'save','original_name':'','site':portal}).status_code,401)
        with client.session_transaction() as session:
            session['admin']=True
            session['csrf']='csrf-test'
        payload={'action':'save','original_name':'','site':portal}
        self.assertEqual(client.post('/api/settings/website',json=payload).status_code,400)
        headers={'X-CSRF-Token':'csrf-test'}
        result=client.post('/api/settings/website',json=payload,headers=headers)
        self.assertEqual(result.status_code,200,result.get_data(as_text=True))
        self.assertEqual(result.json['listener_update'],'unmanaged')
        second=self.site_definition('Downloads','/downloads')
        result=client.post('/api/settings/website',json={'action':'save','original_name':'','site':second},headers=headers)
        self.assertEqual(result.status_code,200,result.get_data(as_text=True))
        self.assertEqual([s.name for s in load_settings(self.data).websites],['Portal','Downloads'])
        # Existing site updates must not overwrite another persisted card.
        changed={**portal,'path':'/apps-v2'}
        result=client.post('/api/settings/website',json={'action':'save','original_name':'Portal','site':changed},headers=headers)
        self.assertEqual(result.status_code,200,result.get_data(as_text=True))
        self.assertEqual([s.path for s in load_settings(self.data).websites],['/apps-v2','/downloads'])
        result=client.post('/api/settings/website',json={'action':'remove','original_name':'Portal','site':None},headers=headers)
        self.assertEqual(result.status_code,200,result.get_data(as_text=True))
        self.assertEqual([s.name for s in load_settings(self.data).websites],['Downloads'])
        self.assertEqual(client.post('/api/settings/website',json={'action':'remove','original_name':'NoSuch','site':None},headers=headers).status_code,400)


if __name__=='__main__':
    unittest.main()
