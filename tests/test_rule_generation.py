import unittest
from src.services.rule_generator import RuleGenerator
from src.models.rule_model import RuleModel

class TestRuleGeneration(unittest.TestCase):

    def setUp(self):
        self.rule_generator = RuleGenerator()
        self.rule_model = RuleModel()

    def test_generate_rule_file(self):
        self.rule_model.add_rule("TestRule", "NetClass1", "KeepoutClearance1")
        result = self.rule_generator.generate_rule_file(self.rule_model)
        self.assertTrue(result)
        self.assertTrue(self.rule_generator.rule_file_exists("TestRule.RUL"))

    def test_invalid_rule_generation(self):
        self.rule_model.add_rule("", "NetClass1", "KeepoutClearance1")
        result = self.rule_generator.generate_rule_file(self.rule_model)
        self.assertFalse(result)

    def test_rule_file_content(self):
        self.rule_model.add_rule("TestRule", "NetClass1", "KeepoutClearance1")
        self.rule_generator.generate_rule_file(self.rule_model)
        content = self.rule_generator.read_rule_file("TestRule.RUL")
        self.assertIn("TestRule", content)
        self.assertIn("NetClass1", content)
        self.assertIn("KeepoutClearance1", content)

if __name__ == '__main__':
    unittest.main()