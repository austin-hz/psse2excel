"""Checks on parsed data. Problems are reported, not fixed."""

from typing import List

import pandas as pd


class PSSEDataValidator:
    """Validates parsed PSS/E data for consistency and correctness."""

    @staticmethod
    def validate_bus_data(buses_df: pd.DataFrame) -> List[str]:
        """Validate bus data for common issues."""
        issues = []

        if buses_df.empty:
            return ["No bus data found"]

        # Check for duplicate bus numbers
        if buses_df['BUS_NUMBER'].duplicated().any():
            duplicates = buses_df[buses_df['BUS_NUMBER'].duplicated()]['BUS_NUMBER'].tolist()
            issues.append(f"Duplicate bus numbers found: {duplicates}")

        # Check for invalid voltage levels
        invalid_voltages = buses_df[buses_df['BASE_KV'] <= 0]['BUS_NUMBER'].tolist()
        if invalid_voltages:
            issues.append(f"Buses with invalid voltage levels: {invalid_voltages}")

        # Check for invalid bus types
        valid_types = [1, 2, 3, 4]
        invalid_types = buses_df[~buses_df['BUS_TYPE'].isin(valid_types)]['BUS_NUMBER'].tolist()
        if invalid_types:
            issues.append(f"Buses with invalid types: {invalid_types}")

        # Check voltage magnitude limits
        invalid_vm = buses_df[(buses_df['VM'] <= 0) | (buses_df['VM'] > 2.0)]['BUS_NUMBER'].tolist()
        if invalid_vm:
            issues.append(f"Buses with invalid voltage magnitudes: {invalid_vm}")

        # Bus angles beyond +/-180 degrees are valid in PSS/E, so they are not checked.

        return issues

    @staticmethod
    def validate_generator_data(generators_df: pd.DataFrame, buses_df: pd.DataFrame) -> List[str]:
        """Validate generator data."""
        issues = []

        if generators_df.empty:
            return issues

        # Check if all generator buses exist in bus data
        if not buses_df.empty:
            missing_buses = set(generators_df['BUS_NUMBER']) - set(buses_df['BUS_NUMBER'])
            if missing_buses:
                issues.append(f"Generators connected to non-existent buses: {list(missing_buses)}")

        # Check for negative generation limits
        negative_pmax = generators_df[generators_df['PMAX'] < 0]['BUS_NUMBER'].tolist()
        if negative_pmax:
            issues.append(f"Generators with negative PMAX: {negative_pmax}")

        # Check for invalid reactive power limits
        invalid_q_limits = generators_df[generators_df['QMAX'] < generators_df['QMIN']]['BUS_NUMBER'].tolist()
        if invalid_q_limits:
            issues.append(f"Generators with QMAX < QMIN: {invalid_q_limits}")

        # Check for unreasonable voltage setpoints
        invalid_vs = generators_df[(generators_df['VS'] < 0.5) | (generators_df['VS'] > 1.5)]['BUS_NUMBER'].tolist()
        if invalid_vs:
            issues.append(f"Generators with unreasonable voltage setpoints: {invalid_vs}")

        # Check for invalid machine base values
        invalid_mbase = generators_df[generators_df['MBASE'] <= 0]['BUS_NUMBER'].tolist()
        if invalid_mbase:
            issues.append(f"Generators with invalid MBASE: {invalid_mbase}")

        return issues

    @staticmethod
    def validate_branch_data(branches_df: pd.DataFrame, buses_df: pd.DataFrame) -> List[str]:
        """Validate branch data."""
        issues = []

        if branches_df.empty:
            return issues

        if not buses_df.empty:
            # Check if all branch buses exist
            all_branch_buses = set(branches_df['FROM_BUS']) | set(branches_df['TO_BUS'])
            missing_buses = all_branch_buses - set(buses_df['BUS_NUMBER'])
            if missing_buses:
                issues.append(f"Branches connected to non-existent buses: {list(missing_buses)}")

        # Check for zero impedance branches (potential issues)
        zero_impedance = branches_df[(branches_df['R'] == 0) & (branches_df['X'] == 0)]
        if not zero_impedance.empty:
            issues.append(f"Branches with zero impedance: {len(zero_impedance)} branches")

        # Check for branches with same from and to bus
        self_loops = branches_df[branches_df['FROM_BUS'] == branches_df['TO_BUS']]
        if not self_loops.empty:
            issues.append(f"Self-loop branches found: {len(self_loops)} branches")

        # Check for negative reactance (unusual but may be valid for some devices)
        negative_x = branches_df[branches_df['X'] < 0]
        if not negative_x.empty:
            issues.append(f"Branches with negative reactance: {len(negative_x)} branches (check if intended)")

        return issues

    @staticmethod
    def validate_load_data(loads_df: pd.DataFrame, buses_df: pd.DataFrame) -> List[str]:
        """Validate load data."""
        issues = []

        if loads_df.empty:
            return issues

        if not buses_df.empty:
            # Check if all load buses exist
            missing_buses = set(loads_df['BUS_NUMBER']) - set(buses_df['BUS_NUMBER'])
            if missing_buses:
                issues.append(f"Loads connected to non-existent buses: {list(missing_buses)}")

        # Check for unreasonably large loads
        large_loads = loads_df[loads_df['PL'] > 10000]['BUS_NUMBER'].tolist()  # 10 GW threshold
        if large_loads:
            issues.append(f"Unreasonably large loads (>10GW): {large_loads}")

        # Check for negative loads (may be valid for generation-like loads)
        negative_loads = loads_df[loads_df['PL'] < 0]['BUS_NUMBER'].tolist()
        if negative_loads:
            issues.append(f"Negative real power loads: {negative_loads} (check if intended)")

        return issues

    @staticmethod
    def validate_transformer_data(transformers_df: pd.DataFrame, buses_df: pd.DataFrame) -> List[str]:
        """Validate transformer data."""
        issues = []

        if transformers_df.empty:
            return issues

        if not buses_df.empty:
            # Check if all transformer buses exist
            all_xfmr_buses = set(transformers_df['FROM_BUS']) | set(transformers_df['TO_BUS'])

            # For 3-winding transformers, also check third bus
            if 'K' in transformers_df.columns:
                three_w_buses = set(transformers_df[transformers_df['K'] != 0]['K'])
                all_xfmr_buses.update(three_w_buses)

            missing_buses = all_xfmr_buses - set(buses_df['BUS_NUMBER'])
            if missing_buses:
                issues.append(f"Transformers connected to non-existent buses: {list(missing_buses)}")

        # Check for zero impedance transformers
        if 'X1_2' in transformers_df.columns:
            zero_impedance = transformers_df[transformers_df['X1_2'] == 0]
            if not zero_impedance.empty:
                issues.append(f"Transformers with zero reactance: {len(zero_impedance)} transformers")

        return issues
