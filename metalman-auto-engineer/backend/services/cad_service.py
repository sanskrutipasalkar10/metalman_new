import cadquery as cq
import os

# Tessellation tolerances tuned for mm-scale sheet-metal/bracket assemblies.
# Too coarse (cadquery's un-tuned default) can flatten small features on thin
# panels into a near-planar mesh; these values keep curved/perforated detail
# without producing an unreasonably large STL.
STL_LINEAR_TOLERANCE = 0.05   # mm
STL_ANGULAR_TOLERANCE = 0.15  # radians

def export_step_to_stl(step_path: str, output_path: str):
    """
    Converts a STEP file to STL for browser-based 3D visualization.
    """
    try:
        print(f"Converting {step_path} to STL...")
        # Import STEP (importStep already combines multi-solid assemblies
        # into a single compound, so the whole assembly is exported, not
        # just the first body).
        result = cq.importers.importStep(step_path)
        # Export STL with explicit tessellation tolerance
        cq.exporters.export(
            result,
            output_path,
            tolerance=STL_LINEAR_TOLERANCE,
            angularTolerance=STL_ANGULAR_TOLERANCE,
        )

        # Log basic mesh stats so a bad/degenerate export is diagnosable
        # from server logs instead of only visible as a "flat" 3D viewer.
        try:
            bbox = result.val().BoundingBox()
            size = (bbox.xlen, bbox.ylen, bbox.zlen)
            stl_size_kb = os.path.getsize(output_path) / 1024
            print(
                f"STL exported to {output_path} "
                f"(bbox mm: {size[0]:.1f} x {size[1]:.1f} x {size[2]:.1f}, "
                f"file size: {stl_size_kb:.1f} KB)"
            )
            if min(size) < 0.01:
                print(
                    "WARNING: exported geometry is near-zero thickness on one axis "
                    "- 3D viewer may render it as a flat plane."
                )
        except Exception as stats_err:
            print(f"STL exported to {output_path} (stats unavailable: {stats_err})")

        return True
    except Exception as e:
        print(f"Error converting STEP to STL: {e}")
        return False
