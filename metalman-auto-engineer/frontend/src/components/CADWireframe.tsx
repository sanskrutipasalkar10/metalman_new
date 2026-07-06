import { Canvas, useFrame, useThree, useLoader } from "@react-three/fiber";
import { Suspense, useRef, useMemo, useEffect, useState } from "react";
import * as THREE from "three";
import { STLLoader } from "three-stdlib";
import { OrbitControls } from "@react-three/drei";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Fits the camera so the model fills ~85 % of the canvas height.
 * Called once after the geometry has been centered & oriented.
 */
const CameraFitter = ({ geom }: { geom: THREE.BufferGeometry }) => {
  const { camera } = useThree();
  useEffect(() => {
    if (!geom) return;
    geom.computeBoundingBox();
    const box = geom.boundingBox;
    if (!box) return;

    const size = new THREE.Vector3();
    box.getSize(size);
    const maxDim = Math.max(size.x, size.y, size.z);

    // FOV-aware distance so the model fills the frame
    const fov = (camera as THREE.PerspectiveCamera).fov * (Math.PI / 180);
    const distance = (maxDim / 2 / Math.tan(fov / 2)) * 1.15; // 1.15 = slight padding

    camera.position.set(0, 0, distance);
    camera.near = distance * 0.01;
    camera.far = distance * 10;
    camera.updateProjectionMatrix();
  }, [geom, camera]);

  return null;
};

// ---------------------------------------------------------------------------
// STL Model
// ---------------------------------------------------------------------------

const STLModel = ({ url }: { url: string }) => {
  const geom = useLoader(STLLoader, url);
  const groupRef = useRef<THREE.Group>(null);

  // Orient geometry: center it, then auto-detect the dominant axis and rotate so
  // the model's largest dimension aligns with Y (upright/portrait on screen).
  // Must run synchronously (useMemo) so CameraFitter sees the final geometry.
  useMemo(() => {
    if (!geom || geom.userData.__oriented) return;

    geom.computeBoundingBox();
    const box = geom.boundingBox;
    const isValid =
      box &&
      Number.isFinite(box.min.x) &&
      Number.isFinite(box.max.x) &&
      !box.isEmpty();

    if (!isValid) {
      console.warn("[CADWireframe] Degenerate bounding box, skipping orientation:", url);
      return;
    }

    // Step 1: Center at origin
    geom.center();

    // Step 2: Detect dominant axis and make it Y (upright on screen)
    geom.computeBoundingBox();
    const size = new THREE.Vector3();
    geom.boundingBox!.getSize(size);

    if (size.x >= size.y && size.x >= size.z) {
      // Model is wider than tall (X dominant) → rotate +90° around Z so X→Y
      geom.applyMatrix4(new THREE.Matrix4().makeRotationZ(Math.PI / 2));
    } else if (size.z >= size.y && size.z >= size.x) {
      // Model is deep (Z dominant, CAD Z-up) → rotate −90° around X so Z→Y
      geom.applyMatrix4(new THREE.Matrix4().makeRotationX(-Math.PI / 2));
    }
    // If Y is already dominant, no rotation needed — model is already upright

    geom.userData.__oriented = true;
  }, [geom, url]);

  // Clockwise rotation when viewed from the front (negative Y-axis)
  useFrame((state) => {
    const t = state.clock.getElapsedTime();
    if (groupRef.current) {
      groupRef.current.rotation.y = -t * 0.4;
    }
  });

  return (
    <>
      <CameraFitter geom={geom} />
      <group ref={groupRef}>
        {/* Solid blue metallic surface */}
        <mesh geometry={geom}>
          <meshStandardMaterial
            color="#1e40af"
            metalness={0.85}
            roughness={0.15}
            emissive="#1d4ed8"
            emissiveIntensity={0.08}
          />
        </mesh>
        {/* Wireframe overlay — scaled 1.003× to sit just above the surface */}
        <mesh geometry={geom} scale={1.003}>
          <meshBasicMaterial color="#60a5fa" wireframe transparent opacity={0.35} />
        </mesh>
      </group>
    </>
  );
};

// ---------------------------------------------------------------------------
// Pulsing ring shown while STL is loading / not yet available
// ---------------------------------------------------------------------------

const LoadingRings = () => {
  const ring1 = useRef<THREE.Mesh>(null);
  const ring2 = useRef<THREE.Mesh>(null);
  const ring3 = useRef<THREE.Mesh>(null);

  useFrame(({ clock }) => {
    const t = clock.getElapsedTime();
    if (ring1.current) ring1.current.rotation.x = t * 0.4;
    if (ring2.current) { ring2.current.rotation.y = t * 0.6; ring2.current.rotation.z = t * 0.2; }
    if (ring3.current) ring3.current.rotation.z = -t * 0.5;
  });

  return (
    <group>
      <mesh ref={ring1}>
        <torusGeometry args={[1.4, 0.04, 16, 80]} />
        <meshBasicMaterial color="#3b82f6" transparent opacity={0.5} />
      </mesh>
      <mesh ref={ring2}>
        <torusGeometry args={[1.0, 0.04, 16, 80]} />
        <meshBasicMaterial color="#60a5fa" transparent opacity={0.4} />
      </mesh>
      <mesh ref={ring3}>
        <torusGeometry args={[0.6, 0.04, 16, 80]} />
        <meshBasicMaterial color="#93c5fd" transparent opacity={0.3} />
      </mesh>
    </group>
  );
};

// ---------------------------------------------------------------------------
// Horizontal scan-plane effect
// ---------------------------------------------------------------------------

const ScanningPlane = () => {
  const ref = useRef<THREE.Mesh>(null);
  useFrame(({ clock }) => {
    const t = clock.getElapsedTime();
    if (ref.current) {
      ref.current.position.y = Math.sin(t * 1.2) * 3;
      if (ref.current.material instanceof THREE.MeshBasicMaterial) {
        ref.current.material.opacity = 0.08 + Math.abs(Math.sin(t * 2.5)) * 0.06;
      }
    }
  });
  return (
    <mesh ref={ref} rotation={[Math.PI / 2, 0, 0]}>
      <planeGeometry args={[20, 20]} />
      <meshBasicMaterial color="#38bdf8" transparent opacity={0.1} side={THREE.DoubleSide} />
    </mesh>
  );
};

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

interface CADWireframeProps {
  stlUrl?: string | null;
}

export const CADWireframe = ({ stlUrl }: CADWireframeProps) => {
  const [stlReady, setStlReady] = useState(false);

  // Reset stlReady flag each time stlUrl changes
  useEffect(() => {
    if (!stlUrl) setStlReady(false);
  }, [stlUrl]);

  return (
    <Canvas
      shadows
      camera={{ position: [0, 0, 8], fov: 45 }}
      dpr={[1, 2]}
      style={{ background: "transparent" }}
    >
      {/* Lighting — no Stage, so we set up lights manually */}
      <ambientLight intensity={0.4} />
      <directionalLight position={[5, 8, 5]} intensity={1.2} castShadow />
      <directionalLight position={[-5, -3, -5]} intensity={0.4} color="#3b82f6" />
      <pointLight position={[0, 4, 0]} intensity={0.6} color="#60a5fa" />

      <Suspense fallback={null}>
        {stlUrl
          ? <STLModel url={stlUrl} />
          : <LoadingRings />  /* Clean pulsing rings — no random geometry */
        }
        <ScanningPlane />
      </Suspense>
    </Canvas>
  );
};
