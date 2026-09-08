import unittest
import os
import sqlite3
import json
from knowledge import (
    init_knowledge_db,
    save_learned_findings,
    get_relevant_knowledge,
    get_all_knowledge,
    add_custom_knowledge_item,
    delete_knowledge_item
)
from app import app


class KnowledgeBaseTests(unittest.TestCase):

    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True
        init_knowledge_db()

    def test_init_db_seeds_initial_findings(self):
        items = get_all_knowledge()
        self.assertGreaterEqual(len(items), 4)

    def test_save_learned_findings_saves_new_finding(self):
        new_finding = [{
            "category": "Inconsistencia Test",
            "title": "Diferencia de prueba en liquidación",
            "situation": "Se constató una inconsistencia relevante en la liquidación de comisión de préstamos.",
            "risk": "Posible sobrepago no detectado.",
            "proposal": "Regularizar la liquidación contable."
        }]
        saved_count = save_learned_findings(new_finding, "Auditoría Préstamos Test")
        self.assertEqual(saved_count, 1)

        relevant = get_relevant_knowledge("inconsistencia comisión liquidación", limit=1)
        self.assertTrue(len(relevant) > 0)
        self.assertIn("liquidación", relevant[0]["situation"].lower())

    def test_knowledge_endpoints(self):
        response = self.app.get("/knowledge")
        self.assertEqual(response.status_code, 200)
        data = json.loads(response.data)
        self.assertIn("items", data)
        self.assertGreaterEqual(data["count"], 4)

        post_res = self.app.post("/knowledge", json={
            "category": "Prueba",
            "title": "Título Regla Prueba",
            "situation": "Situación observada modelo de prueba",
            "risk": "Riesgo modelo",
            "proposal": "Propuesta modelo"
        })
        self.assertEqual(post_res.status_code, 200)
        post_data = json.loads(post_res.data)
        item_id = post_data["id"]

        del_res = self.app.delete(f"/knowledge/{item_id}")
        self.assertEqual(del_res.status_code, 200)


if __name__ == "__main__":
    unittest.main()
