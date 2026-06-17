import unittest

from pydantic import ValidationError

from app.models.schemas import MasterData


class ModelTests(unittest.TestCase):
    def test_master_key_is_trimmed_and_cannot_be_blank(self) -> None:
        self.assertEqual(MasterData(clave=" nit ").clave, "nit")

        with self.assertRaises(ValidationError):
            MasterData(clave="   ")


if __name__ == "__main__":
    unittest.main()
